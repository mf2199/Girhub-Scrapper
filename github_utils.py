"""
Utils for reading data from GitHub issues and building
ready-to-save table with this data.
"""
from github import Github
import operator


# links from labels to project names
PROJECTS = {
    'api: pubsub': 'Pubsub',
    'api: bigtable': 'BigTable',
    'api: bigquery': 'BigQuery',
    'api: firestore': 'FireStore',
    'api: storage': 'Storage',
    'api: spanner': 'Spanner',
    'api: core': 'Core',
}

# repos, from which we track issues
REPO_NAMES = {
    'googleapis/google-cloud-python': 'GCP',
    'googleapis/google-resumable-media-python': 'GRMP'
}

REPO_NAMES_INV = dict((v, k) for k, v in REPO_NAMES.iteritems())
# indexes of field, that must be updated on every update
# other fields will be left unchanged
TRACKED_FIELDS = (1, 2, 6)


def _get_project_name(labels):
    """Designate project name from issue labels."""
    projects = set()

    for label in labels.get_page(0):
        if 'api:' in label.name:
            label = PROJECTS.get(label.name, 'Other')
            projects.add(label)

    projects = sorted(list(projects))
    return ', '.join(projects)


def _build_issue_row(issue, repo):
    """Building one row with issue data."""
    row = []
    if issue.pull_request is None:
        row.append('Medium')
        row.append('=ГИПЕРССЫЛКА("{url}";"{num}")'.format(num=issue.number, url=issue.html_url))
        row.append('PENDING')
        row.append(issue.created_at.strftime('%d %b %Y'))
        row.append(issue.title)
        row.append(REPO_NAMES[repo.full_name])
        row.append(_get_project_name(issue.get_labels()))
        row.append('N/A')
    return row


def build_whole_table():
    """Building ready-to-save table of issues."""
    rows = []
    count = 2

    for repo in REPOS_NAMES.keys():
        repo = gh.get_repo(repo)
        issues = repo.get_issues()

        page_num = 0

        while True:
            issues_page = issues.get_page(page_num)
            if not issues_page:
                break

            for issue in issues_page:
                row = _build_issue_row(issue, repo)
                if row:
                    rows.append(row)
                    count += 1

            page_num += 1

    rows.sort(key=operator.itemgetter(5, 6, 1))
    return rows, count


gh = Github('q-logic', 'qlogic1352')
