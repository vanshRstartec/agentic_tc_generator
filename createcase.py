import requests
from requests.auth import HTTPBasicAuth


class TestCaseManager:
    def __init__(self, org, proj, pat, plan_name):
        self.org = org
        self.proj = proj
        self.base = f"https://dev.azure.com/{org}/{proj}/_apis"
        self.auth = HTTPBasicAuth('', pat)
        self.h = {"Content-Type": "application/json"}
        self.plan_name = plan_name
        self.plan_id = None
        self.suites = {}
        self._setup_plan()

    def _setup_plan(self):
        r = requests.get(f"{self.base}/testplan/plans?api-version=7.0", headers=self.h, auth=self.auth)
        plans = r.json().get('value', [])
        self.plan_id = next((p['id'] for p in plans if p['name'] == self.plan_name), None)

        if not self.plan_id:
            r = requests.post(f"{self.base}/testplan/plans?api-version=7.0", headers=self.h, auth=self.auth,
                              json={"name": self.plan_name, "areaPath": self.proj, "iteration": self.proj})
            self.plan_id = r.json()['id']

    def _get_suite(self, suite_name):
        if suite_name in self.suites:
            return self.suites[suite_name]

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

        payload = [
            {"op": "add", "path": "/fields/System.Title", "value": title},
            {"op": "add", "path": "/fields/Microsoft.VSTS.TCM.Steps", "value": steps_xml},
            {"op": "add", "path": "/fields/Microsoft.VSTS.Common.Priority", "value": priority}
        ]

        r = requests.post(f"{self.base}/wit/workitems/$Test Case?api-version=7.0",
                          headers={"Content-Type": "application/json-patch+json"}, auth=self.auth, json=payload)
        tc_id = r.json()['id']

        requests.post(
            f"https://dev.azure.com/{self.org}/{self.proj}/_apis/test/Plans/{self.plan_id}/Suites/{suite_id}/testcases/{tc_id}?api-version=5.0",
            headers=self.h, auth=self.auth)

        print(f"Created Test Case #{tc_id} in suite '{suite_name}'")
        return tc_id