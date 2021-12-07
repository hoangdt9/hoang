import utils
import requests
import json, sys
from datetime import date, datetime, timedelta

space_key = "SVMC"
# parentTitle = "Project Report - Automatic"
# weeklyPageTitle = "Weekly Project Status Report"
# monthlyPageTitle = "CP Monthly Report"

dailyPageTitle = "Issue Status Tool"
pageUrgentPrjTitle = "Issue Tool Project List"
user = utils.open_file(".plm")[0]
pw = utils.open_file(".plm")[2]


def submitToWiki(page_title, page_content):
    response = getPageContent(page_title, space_key)
    if response.json()['size'] > 0:
        print('update page %s' % page_title)
        page_id = response.json()['results'][0]['id']
        current_version = response.json()['results'][0]['version']['number']

        data = {
            'id': str(page_id),
            'type': 'page',
            'title': page_title,
            'space': {'key': space_key},
            'version': {'number': current_version + 1},
            'body': {
                'storage':
                    {
                        'value': str(page_content),
                        'representation': 'storage',
                    }
            }
        }

        data_to_send = json.dumps(data).encode("utf-8")

        response = requests.put('http://mobilerndhub.sec.samsung.net/wiki/rest/api/content/%s' % page_id,
                                headers={'Content-Type': 'application/json'}, data=data_to_send, auth=(user, pw))

        if response.status_code == requests.codes['ok']:
            print("View page at %s" % response.url)

    else:
        print('add page %s' % page_title)
        response = requests.get('http://mobilerndhub.sec.samsung.net/wiki/rest/api/content?spaceKey=%s&title=%s' %
                                (space_key, parentTitle), auth=(user, pw))
        parent_id = response.json()['results'][0]['id']

        data = {
            'type': 'page',
            'title': page_title,
            "ancestors": [{"id": parent_id}],
            'space': {'key': space_key},
            'body': {
                'storage':
                    {
                        'value': str(page_content),
                        'representation': 'storage',
                    }
            }
        }

        data_to_send = json.dumps(data).encode("utf-8")

        response = requests.post('http://mobilerndhub.sec.samsung.net/wiki/rest/api/content/',
                                 headers={'Content-Type': 'application/json'}, data=data_to_send, auth=(user, pw))

        if response.status_code == requests.codes['ok']:
            print("View page at %s" % response.url)


def getPageContent(pageTitle, space_key):
    response = requests.get('http://mobilerndhub.sec.samsung.net/wiki/rest/api/content?spaceKey=%s&title=%s&'
                            'expand=space,body.view,version,container' % (space_key, pageTitle), auth=(user, pw))
    if not response.status_code == requests.codes['ok']:
        print("Cannot get content of page: " + pageTitle)
        sys.exit(1)

    return response


def getListSingleID(data):
    """
    :param data:  table data
    :return: list mysingle to chart group
    """
    list_id = []
    index = data[0].index('Owner')
    for i in data:
        list_id.append(i[index])
    del (list_id[0])

    return list_id


def makeLinkChat(mySingleId):
    """Returns <a> tag with href from single ID"""
    info_link = "mysingleim://%s"
    return r"<a target='_blank' href='%s'>%s</a>" % (info_link % mySingleId, mySingleId)


def makeLinkNameChat(mySingleId, name_member):
    """Returns <a> tag with href from single ID"""
    info_link = "mysingleim://%s"
    return r"<a target='_blank' href='%s'>%s</a>" % (info_link % mySingleId, name_member)


def makeLinkChatGroup(listID):
    """Returns <a> tag with href from single ID"""
    strListID = ""
    for i in range(0, len(listID)):
        strListID += str(listID[i]) + ';'
    info_link = "mysingleim://%s"
    return r"<a target='_blank' style='font-size: 12px; font-style: normal;' target='_blank'  href='%s'>%s</a>" % (
        info_link % strListID, "<br />Chat")


def makeLinkPLM(PLMCaseCode):
    """Returns <a> tag with href from mysingleID"""
    return "<a target='_blank' href='http://splm.sec.samsung.net/wl/tqm/defect/defectreg/getDefectCodeSearch.do?defectCode=%s'>%s</a>" % (
    PLMCaseCode, PLMCaseCode)


def make_link_chat(single_id, text):
    """Returns <a> tag with href from single ID"""
    info_link = "mysingleim://%s"
    return r"<a target='_blank' href='%s'>%s</a>" % (info_link % single_id, text)


def make_link_jira(jira_key):
    jira_link = r"http://mobilerndhub.sec.samsung.net/its/browse/%s"
    return r"<a target='_blank' href='%s'>%s</a>" % (jira_link % jira_key, jira_key)


def make_link_jira_with_summary(jira_key, text):
    jira_link = r"http://mobilerndhub.sec.samsung.net/its/browse/%s"
    return r"<a target='_blank' href='%s'>%s</a>" % (jira_link % jira_key, text)


def make_img_jira(link):
    return r"<img src='%s' class='icon'>" % link


def make_status_jira(text):
    if text.lower() == 'new':
        return r"<span class='aui-lozenge aui-lozenge-subtle aui-lozenge-complete'>%s</span>" % text
    else:
        return r"<span class='aui-lozenge aui-lozenge-subtle aui-lozenge-current'>%s</span>" % text


def create_isssue_owner(owner_list):
    html = "<head> \n </head> \n <body> \n <div> \n <p>"

    for i in owner_list:
        key = get_user_key(i)
        html += '<ac:link><ri:user ri:userkey="%s" /></ac:link>' % key
        html += ", "

    html += "</p> \n </div> \n </body>"
    return html


def check_time_update():
    response = getPageContent(dailyPageTitle, space_key)
    page_key = response.json()['results'][0]['id']
    response = requests.get("http://mobilerndhub.sec.samsung.net/wiki/rest/api/content/%s/history" % str(page_key),
                            auth=(user, pw))
    time_update = response.json()['lastUpdated']['when'][:19]  # %Y-%m-%dT%H:%M:%S
    datetime_update = datetime.strptime(time_update, "%Y-%m-%dT%H:%M:%S") - timedelta(hours=2)  # HQ earlier VN 2 hours
    print("latest time update page: %s" % datetime_update.strftime("%H:%M %d-%m-%Y"))
    return datetime_update


def get_updated_date(pageTitle):
    response = getPageContent(pageTitle, space_key)
    page_key = response.json()['results'][0]['id']
    response = requests.get("http://mobilerndhub.sec.samsung.net/wiki/rest/api/content/%s/history" % str(page_key),
                            auth=(user, pw))
    return response.json()['lastUpdated']['when'][:10]  # YYYY-MM-DD


def get_user_key(user_name):
    request_data = requests.get("http://mobilerndhub.sec.samsung.net/wiki/rest/api/user?username=%s" % user_name,
                                auth=(user, pw))
    return request_data.json()['userKey']


def get_all_data_jira_task_list(project_key):
    # Query data with in 3 month
    jql_query = "project = %s and status not in (resolved, cancelled) and created > startOfMonth(-2) order by " \
                "created desc" % project_key
    max_result = 1000

    params = {
        "jql": jql_query,
        "startAt": 0,
        "maxResults": max_result,
        "fields": [
            "key",
            "summary",
            "issuetype",
            "created",
            "duedate",
            "resolutiondate",
            "assignee",
            "priority",
            "status"
        ]
    }

    url_query = 'http://mobilerndhub.sec.samsung.net/its/rest/api/2/search'
    data_task_list_json = requests.get(url_query, params=params, auth=(user, pw))

    list_all_task = json.loads(data_task_list_json.text)
    return list_all_task['issues']


def convert_date_time(date_time):
    date_time = datetime.strptime(date_time, "%Y-%m-%d").date()
    return date_time


def get_data_jira_task_list_by_team(all_data_jira_task_list, member_id_list):
    num_of_jira_task_by_team = {}
    info_detail_jira_task = []
    data_jira_task_for_pie_chart = [["", 'Jira Tasks'], ['Done', 0], ['NEW', 0], ["In Progress", 0]]

    list_all_member = []
    for team, member_of_team in member_id_list.items():
        num_of_jira_task_by_team[team] = [0, 0]  # [open, in progress]
        list_all_member += member_of_team
    number_of_jira_task_by_member = {key: 0 for key in list_all_member}

    for task_info in all_data_jira_task_list:
        summary = task_info['fields']['summary']

        if not summary.startswith('[Automatic]'):
            due_date = task_info['fields']['duedate']
            created = task_info['fields']['created'][:10]
            resolve_date = task_info['fields']['resolutiondate']

            if resolve_date is None:
                resolve_date = ''
            else:
                resolve_date = convert_date_time(resolve_date[:10])

            if due_date is None:
                due_date = ''
            # else:
            #     due_date = convert_date_time(due_date)

            single_id = task_info['fields']['assignee']['key']
            team = ""

            status_jira = task_info['fields']['status']['name'].lower()

            if status_jira == 'in progress':
                data_jira_task_for_pie_chart[3][1] += 1
            elif status_jira == 'new':
                data_jira_task_for_pie_chart[2][1] += 1
            else:
                data_jira_task_for_pie_chart[1][1] += 1

            if status_jira == 'done' and resolve_date == date.today():
                # include jira task resolve to day
                number_of_jira_task_by_member[single_id] += 1
            if status_jira == 'in progress' or status_jira == 'new':
                try:
                    number_of_jira_task_by_member[single_id] += 1
                except KeyError:
                    number_of_jira_task_by_member[single_id] = 1

                for key, value in member_id_list.items():
                    if single_id in value:
                        team = key
                        if status_jira == 'in progress':
                            num_of_jira_task_by_team[key][1] = num_of_jira_task_by_team[key][1] + 1
                        elif status_jira == 'new':
                            num_of_jira_task_by_team[key][0] = num_of_jira_task_by_team[key][0] + 1
                        break

                info = [
                    make_link_jira(task_info['key']),
                    summary,
                    make_img_jira(task_info['fields']['issuetype']['iconUrl']),
                    created,
                    due_date,
                    make_link_chat(single_id, task_info['fields']['assignee']['displayName']),
                    team,
                    make_img_jira(task_info['fields']['priority']['iconUrl']),
                    make_status_jira(task_info['fields']['status']['name'])
                ]

                info_detail_jira_task.append(info)

    data_chart_pie_jira = 'var dataChartPieJira = ' + str(data_jira_task_for_pie_chart) + '; \n'

    return num_of_jira_task_by_team, info_detail_jira_task, number_of_jira_task_by_member, data_chart_pie_jira
