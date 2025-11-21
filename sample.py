import openai
import pandas as pd
import os
import ast
from dotenv import load_dotenv
from createcase import TestCaseManager

load_dotenv()
openai.api_type = "azure"
openai.api_base = os.getenv("AZURE_OPENAI_ENDPOINT")
openai.api_version = "2024-02-15-preview"
openai.api_key = os.getenv("AZURE_OPENAI_KEY")


def parse_steps_to_list(steps_str):
    try:
        steps_str = steps_str.strip()
        if not steps_str.startswith('['): steps_str = '[' + steps_str + ']'
        steps_list = ast.literal_eval(steps_str)
        if isinstance(steps_list, list) and all(
                isinstance(s, dict) and 'action' in s and 'expected' in s for s in steps_list):
            return steps_list
    except:
        pass
    return []


def generate_test_cases(input_file, output_file="generated_test_cases.xlsx", mgr=None, suite_name="LOGIN"):
    df = pd.read_excel(input_file)
    output, current_story = [], ""

    for _, row in df.iterrows():
        user_story = row.get("User Story", "")
        user_story = user_story if pd.notna(user_story) else current_story
        current_story = user_story
        ac = row.get("Acceptance Criteria", "")

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
            response = openai.ChatCompletion.create(
                engine=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
                messages=[{"role": "system",
                           "content": "You are a QA engineer. Generate test cases in the EXACT format requested."},
                          {"role": "user", "content": prompt}],
                temperature=0.3, max_tokens=1500
            )

            content = response.choices[0].message.content.strip()
            blocks = [b for b in content.split("---") if "Title:" in b]

            for block in blocks:
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
                    if line.strip().startswith("```"):
                        in_code_block = not in_code_block
                        continue
                    if in_code_block: steps += line + "\n"

                output.append(
                    {"User Story": user_story, "Acceptance Criteria": ac, "Test Type": test_type, "Title": title,
                     "Priority": priority, "Steps": steps.strip(), "Status": "Not Executed", "Comments": ""})

        except Exception as e:
            output.append({"User Story": user_story, "Acceptance Criteria": ac, "Test Type": "Error",
                           "Title": "Generation Failed", "Priority": 2, "Steps": "N/A", "Status": "Error",
                           "Comments": str(e)})

    df_out = pd.DataFrame(output)
    df_out.insert(0, "S.No.", range(1, len(df_out) + 1))
    df_out.to_excel(output_file, index=False)
    print(f"✅ Generated {len(df_out)} test cases → {output_file}")

    if mgr:
        print(f"🔄 Uploading to ADO suite '{suite_name}'...")
        upload_count = error_count = 0

        for _, row in df_out.iterrows():
            if row["Status"] == "Error": continue
            steps_list = parse_steps_to_list(row["Steps"])
            if not steps_list:
                error_count += 1
                continue

            try:
                mgr.create_test_case(suite_name=suite_name, title=row["Title"], steps=steps_list,
                                     priority=row["Priority"])
                upload_count += 1
            except Exception as e:
                error_count += 1
                print(f"❌ {row['Title']}: {str(e)}")

        print(f"✅ Uploaded {upload_count}/{len(df_out)} test cases ({error_count} failed)")

    return df_out