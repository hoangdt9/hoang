import ReadDataFromExcel as DataEXL


def create_data_urgent(type_data="URGENT"):
    """ create variable for chart urgent """
    if type_data == "URGENT":
        input_data = DataEXL.urgent_count_all_prj()
    else:
        input_data = DataEXL.flagship_count_all_prj()

    list_team = DataEXL.get_list_team()
    total_data = dict((team, [0, 0]) for team in list_team)

    for prj in input_data:
        for team in input_data[prj]:
            total_data[team][0] += input_data[prj][team][0]
            total_data[team][1] += input_data[prj][team][1]

    prj_detail = {}
    for prj in input_data:
        prj_detail[prj] = convert_data(input_data[prj])

    if type_data == "URGENT":
        out_total_data = "var all_models = " + str(convert_data(total_data)) + ";\n"
        out_detail_data = "var prj_urgent_detail = " + str(prj_detail) + '; \n'
    else:
        out_total_data = "var all_flagship_models = " + str(convert_data(total_data)) + ";\n"
        out_detail_data = "var prj_flagship_detail = " + str(prj_detail) + '; \n'
    data_chart = out_total_data + out_detail_data

    return data_chart


def create_data_plm_jira_issue(main_summary_data, num_of_jira_task_by_team):
    list_summary_plm = []

    for key, value in num_of_jira_task_by_team.items():
        num_of_task_jira = value[0] + value[1]
        num_of_issue_plm = main_summary_data[key]

        plm_issue = [key, num_of_issue_plm, num_of_task_jira]
        list_summary_plm.append(plm_issue)

    # Sort by team name A -> Z
    list_summary_plm.sort(key=lambda item : item[0])
    data_summary_plm = "var summary_data_chart = " + str([['', 'PLM Issue', 'Jira Task']] + list_summary_plm) + "; \n"
    return data_summary_plm


def convert_data(dict_data):
    """ convert dict data to chart """
    out_data = [["", "Today Resolved", "Open"]]
    counter = 0
    for team in dict_data:
        if dict_data[team][0] > 0 or dict_data[team][1] > 0:
            counter += 1
            out_data.append([team, dict_data[team][1], dict_data[team][0]])  # [team name, resolved, open]
    if counter == 1:
        out_data.append([" ", 0, 0])
    return out_data
