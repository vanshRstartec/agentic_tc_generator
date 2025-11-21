import openai
import pandas as pd
import os
import ast
import requests
from requests.auth import HTTPBasicAuth
from dotenv import load_dotenv

load_dotenv()
openai.api_type = "azure"
openai.api_base = os.getenv("AZURE_OPENAI_ENDPOINT")
openai.api_version = "2024-02-15-preview"
openai.api_key = os.getenv("AZURE_OPENAI_KEY")


class ADOTestManager:
    def __init__(self, org, proj, pat, plan_name):
        self.org, self.proj, self.base = org, proj, f"https://dev.azure.com/{org}/{proj}/_apis"
        self.auth, self.h, self.suites = HTTPBasicAuth('', pat), {"Content-Type": "application/json"}, {}
        self.plan_id = self._setup_plan(plan_name)

    def _setup_plan(self, plan_name):
        r = requests.get(f"{self.base}/testplan/plans?api-version=7.0", headers=self.h, auth=self.auth)
        plan_id = next((p['id'] for p in r.json().get('value', []) if p['name'] == plan_name), None)
        if not plan_id:
            r = requests.post(f"{self.base}/testplan/plans?api-version=7.0", headers=self.h, auth=self.auth,
                              json={"name": plan_name, "areaPath": self.proj, "iteration": self.proj})
            plan_id = r.json()['id']
        return plan_id

    def _get_suite(self, suite_name):
        if suite_name in self.suites: return self.suites[suite_name]
        r = requests.get(f"{self.base}/testplan/plans/{self.plan_id}/suites?api-version=7.0", headers=self.h,
                         auth=self.auth)
        suites = r.json().get('value', [])
        root = next((s for s in suites if s.get('suiteType') == 'staticTestSuite' and s.get('parentSuite') is None),
                    suites[0])
        suite_id = next((s['id'] for s in suites if s['name'] == suite_name), None)
        if not suite_id:
            r = requests.post(f"{self.base}/testplan/plans/{self.plan_id}/suites?api-version=7.0", headers=self.h,
                              auth=self.auth,
                              json={"suiteType": "staticTestSuite", "name": suite_name,
                                    "parentSuite": {"id": root['id']}})
            suite_id = r.json()['id']
        self.suites[suite_name] = suite_id
        return suite_id

    def create_test_case(self, suite_name, title, steps, priority=2):
        suite_id = self._get_suite(suite_name)
        steps_xml = f'<steps id="0" last="{len(steps)}">'
        for i, s in enumerate(steps, 1):
            steps_xml += f'<step id="{i}" type="ActionStep"><parameterizedString isformatted="true">&lt;P&gt;{s["action"]}&lt;/P&gt;</parameterizedString><parameterizedString isformatted="true">&lt;P&gt;{s["expected"]}&lt;/P&gt;</parameterizedString><description/></step>'
        steps_xml += '</steps>'
        payload = [{"op": "add", "path": "/fields/System.Title", "value": title},
                   {"op": "add", "path": "/fields/Microsoft.VSTS.TCM.Steps", "value": steps_xml},
                   {"op": "add", "path": "/fields/Microsoft.VSTS.Common.Priority", "value": priority}]
        r = requests.post(f"{self.base}/wit/workitems/$Test Case?api-version=7.0",
                          headers={"Content-Type": "application/json-patch+json"}, auth=self.auth, json=payload)
        tc_id = r.json()['id']
        requests.post(
            f"https://dev.azure.com/{self.org}/{self.proj}/_apis/test/Plans/{self.plan_id}/Suites/{suite_id}/testcases/{tc_id}?api-version=5.0",
            headers=self.h, auth=self.auth)
        print(f"Created Test Case #{tc_id} in suite '{suite_name}'")
        return tc_id


def generate_test_cases(input_file, output_file=None):
    from datetime import datetime
    if output_file is None:
        os.makedirs("output", exist_ok=True)
        output_file = f"output/{datetime.now().strftime('%Y%m%d_%H%M%S')}_generated_tcs.xlsx"
    df, output, current_story = pd.read_excel(input_file), [], ""
    if output_dir := os.path.dirname(output_file): os.makedirs(output_dir, exist_ok=True)
    for _, row in df.iterrows():
        user_story = row.get("User Story", "") if pd.notna(row.get("User Story", "")) else current_story
        current_story, ac = user_story, row.get("Acceptance Criteria", "")
        context = "\n".join([f"{col}: {row.get(col, '')}" for col in
                             ["Feature/Module", "Priority", "Risk Level", "Preconditions", "Test Environment",
                              "Generic Test Data", "Comments/Notes"] if
                             pd.notna(row.get(col, "")) and str(row.get(col, "")).strip()])
        prompt = f"""Generate test cases for:
User Story: {user_story}
Acceptance Criteria: {ac}
{context}

Create 3 test cases (Positive, Negative, Edge). Format each EXACTLY as:

Test Type: [Positive/Negative/Edge]
Title: [Clear test case title]
Priority: [1-4]
Steps:
```
{{'action': '[step action]', 'expected': '[step expected]'}},
{{'action': '[step action]', 'expected': '[step expected]'}},
{{'action': '[step action]', 'expected': '[step expected]'}}
```
---
Each step must be a dictionary with 'action' and 'expected' keys."""
        try:
            response = openai.ChatCompletion.create(engine=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
                                                    messages=[{"role": "system",
                                                               "content": "You are a QA engineer. Generate test cases in the EXACT format requested."},
                                                              {"role": "user", "content": prompt}], temperature=0.3,
                                                    max_tokens=1500)
            content = response.choices[0].message.content.strip()
            for block in [b for b in content.split("---") if "Title:" in b]:
                lines = block.strip().split("\n")
                test_type = next((l.replace("Test Type:", "").strip() for l in lines if "Test Type:" in l), "Unknown")
                title = next((l.replace("Title:", "").strip() for l in lines if "Title:" in l), "Test Case")
                priority = next((l.replace("Priority:", "").strip() for l in lines if "Priority:" in l), "2")
                try:
                    priority = int(priority)
                    if priority not in [1, 2, 3, 4]: priority = 2
                except:
                    priority = 2
                steps, in_code_block = "", False
                for line in lines:
                    if line.strip().startswith("```"): in_code_block = not in_code_block; continue
                    if in_code_block: steps += line + "\n"

                # Convert steps from dict format to action -> expected format
                steps_formatted = []
                try:
                    if not steps.strip().startswith('['): steps = '[' + steps.strip() + ']'
                    steps_list = ast.literal_eval(steps.strip())
                    if isinstance(steps_list, list):
                        for s in steps_list:
                            if isinstance(s, dict) and 'action' in s and 'expected' in s:
                                steps_formatted.append(f"{s['action']} -> {s['expected']}")
                    steps = '\n'.join(steps_formatted) if steps_formatted else steps.strip()
                except:
                    steps = steps.strip()

                output.append(
                    {"User Story": user_story, "Acceptance Criteria": ac, "Test Type": test_type, "Title": title,
                     "Priority": priority, "Steps": steps, "Status": "Not Executed", "Comments": ""})
        except Exception as e:
            output.append({"User Story": user_story, "Acceptance Criteria": ac, "Test Type": "Error",
                           "Title": "Generation Failed", "Priority": 2, "Steps": "N/A", "Status": "Error",
                           "Comments": str(e)})
    df_out = pd.DataFrame(output)
    df_out.insert(0, "S.No.", range(1, len(df_out) + 1))
    df_out.to_excel(output_file, index=False)
    print(f"✅ Generated {len(df_out)} test cases → {output_file}")
    return output_file


def upload_test_cases(excel_file, org, proj, pat, plan_name, suite_name="LOGIN"):
    df, mgr = pd.read_excel(excel_file), ADOTestManager(org, proj, pat, plan_name)
    upload_count, error_count = 0, 0
    print(f"🔄 Uploading to ADO suite '{suite_name}'...")
    for _, row in df.iterrows():
        if row.get("Status") == "Error": continue
        steps_str = str(row.get("Steps", "")).strip()
        try:
            # Parse steps from "action -> expected" format
            steps_list = []
            for line in steps_str.split('\n'):
                line = line.strip()
                if '->' in line:
                    parts = line.split('->', 1)
                    steps_list.append({'action': parts[0].strip(), 'expected': parts[1].strip()})

            if not steps_list:
                error_count += 1;
                continue

            mgr.create_test_case(suite_name=suite_name, title=row.get("Title", "Test Case"), steps=steps_list,
                                 priority=int(row.get("Priority", 2)))
            upload_count += 1
        except Exception as e:
            error_count += 1
            print(f"❌ {row.get('Title', 'Unknown')}: {str(e)}")
    print(f"✅ Uploaded {upload_count}/{len(df)} test cases ({error_count} failed)")
    return upload_count, error_count
