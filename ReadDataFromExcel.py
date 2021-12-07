import xlwings as xw
import pandas as pd
import re
import utils
import pythoncom
from datetime import date, datetime, timedelta
import WikiSubmit as Wiki
import V2_UrgentTable
from TabTeamDetail import DataTabTeamDetail

LIST_MEMBER = r'sample input\\Member_Organization.xlsx'
LIST_ISSUE = r'sample input\\DEFECT_LIST_Today_Basic.xls'

list_urgent = []
list_team = []
list_tg = []
list_team_of_tg = {}
member_id_by_team = {}
list_id_and_name_member = {}
urgent_issue_prj_count = {}
flagship_issue_prj_count = {}
dfExcel = pd.DataFrame()

owner_issue_urgent = []
owner_issue_pending = []


def get_urgent_project():
    """ get list urgent project """
    return V2_UrgentTable.get_list_project('Urgent')


def get_flagship_prject():
    """ get list flagship project """
    return V2_UrgentTable.get_list_project('Flagship')


def team_member(id_member, member_id_by_team):
    for key, value in member_id_by_team.items():
        if id_member in value:
            return key


def is_pending_issue(ppt, time_pending_standard):
    """ check is long pending issue """
    int_list = [int(s) for s in re.findall('\\d+', ppt)]
    if int_list[0] > time_pending_standard:
        return True
    elif int_list[0] == time_pending_standard and (int_list[1] > 0 or int_list[2] > 0):
        return True
    else:
        return False


def ppt_min(issue):
    """ get time pending issue """
    ppt_index = 2
    int_list = [int(s) for s in re.findall('\\d+', issue[ppt_index])]
    return int_list[0] * 24 * 60 + int_list[1] * 60 + int_list[2]


def get_list_single_id_member():
    """ return list single id all mem of Part """
    pythoncom.CoInitialize()
    try:
        xw.App(visible=False)
        memberBook = xw.Book(LIST_MEMBER)
        memberSheet = memberBook.sheets("My Organization Member")
        tableValue = memberSheet.range('A2').expand('table').value
        length = len(memberSheet.range('A2').expand('right').value)

        dataConvert = pd.DataFrame(tableValue, columns=memberSheet.range('A1').expand('right').value[:length])
        list_single_id_member = dataConvert["mySingle"].tolist()
        list_single_id_member = [single_id.lower() for single_id in list_single_id_member]
        memberBook.close()
        return list_single_id_member

    except Exception as e:
        print(e)
    finally:
        pythoncom.CoUninitialize()


def get_team_info():
    """ get total info team of part """
    pythoncom.CoInitialize()
    try:
        xw.App(visible=False)
        memberBook = xw.Book(LIST_MEMBER)
        memberSheet = memberBook.sheets("My Organization Member")
        tableValue = memberSheet.range('A2').expand('table').value
        length = len(memberSheet.range('A2').expand('right').value)

        dataConvert = pd.DataFrame(tableValue, columns=memberSheet.range('A1').expand('right').value[:length])
        list_tg = list(set(dataConvert["SVMC TG"].tolist()))
        list_team = list(set(dataConvert["SVMC Sub TG"].tolist()))

        list_single_id_member = dataConvert["mySingle"].tolist()
        list_single_id_member = [single_id.lower() for single_id in list_single_id_member]
        list_name_member = dataConvert["Full Name"].tolist()

        # remove all space start-end element
        list_tg = list(map(str.strip, list_tg))
        list_team = list(map(str.strip, list_team))
        list_team_name = []

        all_member_id_part = dataConvert["mySingle"].tolist()
        numOfMember = len(all_member_id_part)

        list_team_of_tg = {}
        member_name_by_team = {}
        member_id_by_team = {}
        list_id_and_name_member = dict(zip(list_single_id_member, list_name_member))
        # create list team of TG
        for tg in list_tg:
            dataTG = dataConvert.loc[dataConvert["SVMC TG"] == tg]
            teamList = dataTG["SVMC Sub TG"].tolist()
            teamList = [team_name.replace(' &amp; ', '_').replace(' & ', '_').replace(' ', '_').replace('/', '_')
                        for team_name in list(set(teamList))]
            list_team_of_tg[tg] = teamList

        # create list member Name/ID of Team
        for team in list_team:
            team_name = team.replace(' &amp; ', '_').replace(' ', '_').replace('/', '_')
            list_team_name.append(team_name)
            dataTeam = dataConvert.loc[dataConvert["SVMC Sub TG"] == team]

            list_member_name = dataTeam["Full Name"].tolist()
            member_of_team = [name.title() for name in
                              list_member_name]  # name = proper case (John Smith) , id = lower case
            member_name_by_team[team_name] = member_of_team
            team_member_id = dataTeam["mySingle"].tolist()
            team_member_id = [id.lower() for id in team_member_id]
            member_id_by_team[team_name] = team_member_id

        memberBook.close()
        return list_tg, list_team_of_tg, list_team_name, numOfMember, \
               member_name_by_team, member_id_by_team, all_member_id_part, list_id_and_name_member
    except Exception as e:
        print(e)
    finally:
        pythoncom.CoUninitialize()


def is_urgent_issue(name, urgent_project):
    """ check is urgent issue """
    for prj in urgent_project:
        if prj in name:
            return prj
    return ''


def folder_issue(data, index):
    """ return to folder of issue """
    try:
        if data['Occurr. Stg.'][index] in str(['DI', 'PI', 'DV', 'PV', 'PR', 'SR', 'MS']):
            return 'Main'
        elif 'MR' in data['Pjt. Name'][index]:
            return 'MR'
        else:
            return 'Sub'
    except:
        print("error format of Occurr. Stg.")


def summary_analysis(input_data):
    """ convert data to summary report """
    chart_summary = {}
    for team in input_data:
        #todo 04 June 2020 just get open issue
        open_issue = input_data[team][0]  # just get open issue
        chart_summary[team] = open_issue

    return chart_summary


def summary_analysis_active(input_data, type_data):
    """ create data include open-active """
    issue_open = {'type_chart': type_data, 'type_data': 'open'}
    issue_active = {'type_chart': type_data, 'type_data': 'active'}
    for team in input_data:
        active_issue = input_data[team][0] + input_data[team][2]
        issue_open[team] = input_data[team][0]
        issue_active[team] = active_issue

    summary_open_active = [issue_open, issue_active]
    return summary_open_active


def read_excel_file(file_name, sheet_name):
    """ read issue from excel to dataFrame """
    pythoncom.CoInitialize()
    xw.App(visible=False)
    excelBook = xw.Book(file_name)
    excelSheet = excelBook.sheets(sheet_name)
    pd_header = excelSheet.range('A3').expand('right').value
    num_row = len(excelSheet.range('A3').expand('down').value) + 2
    num_column = len(pd_header)
    if num_column <= 26:
        # if column name in range A - Z
        last_column_name = chr(64 + num_column)
    else:
        # if number of column > 26 name column change to AA - AZ
        last_column_name = 'A' + chr(64 + (num_column - 26))

    if num_row < 1000:
        range_table = 'A4:' + last_column_name + str(num_row)
        table_values = excelSheet.range(range_table).value
        data = pd.DataFrame(table_values, columns=pd_header)
    else:
        range_table1 = 'A4:'
        range_table2 = 'A'
        split_point = int(num_row / 2)
        range_table1 += last_column_name + str(split_point)
        range_table2 += str(split_point + 1) + ':' + last_column_name + str(num_row)

        table_value1 = excelSheet.range(range_table1).value
        table_value2 = excelSheet.range(range_table2).value

        df1 = pd.DataFrame(table_value1, columns=pd_header)
        df2 = pd.DataFrame(table_value2, columns=pd_header)
        data = df1.append(df2, ignore_index=True)

    length_frame = len(data)
    if sheet_name != "DEFECT":
        data = data.loc[data['Type'] == 'Comment']
        data = data.drop_duplicates(subset=['Case Code.'], keep='first')
    else:
        data = data.drop_duplicates(subset=['Case Code'], keep='last')
    excelBook.close()
    pythoncom.CoUninitialize()
    # print("number of row of file %s after drop duplicated: %d" % (file_name, len(data)))
    return data, length_frame  # luc.nc: Must return length of frame before drop


def flagship_count_prj(prj):
    """ get total count detail flagship project """
    tempdict = flagship_issue_prj_count.get(prj, {})
    num_open_issue = 0
    num_resolved_issue = 0
    for issues in tempdict.values():
        num_open_issue += issues[0]
        num_resolved_issue += issues[1]
    return num_open_issue, num_resolved_issue


def flagship_count_all_prj():
    return flagship_issue_prj_count


def urgent_count_prj(prj):
    """ get total count detail urgent project """
    tempdict = urgent_issue_prj_count.get(prj, {})
    num_open_issue = 0
    num_resolved_issue = 0
    for issues in tempdict.values():
        num_open_issue += issues[0]
        num_resolved_issue += issues[1]
    return num_open_issue, num_resolved_issue


def urgent_count_all_prj():
    return urgent_issue_prj_count


def get_urgent_owner():
    """ return owner of issue urgent """
    return owner_issue_urgent


def get_data_frame():
    """ return to dataFrame read from excel """
    return dfExcel


def get_list_team_of_tg():
    """ return dict team of TG """
    return list_team_of_tg


def get_member_id_of_team():
    """ return dict member id of Team """
    return member_id_by_team


def get_list_team():
    """ return to team name of part """
    return list_team


def get_list_tg():
    """ return to team name of part """
    return list_tg


def get_list_id_and_name_member():
    """ return to single_id and name member of part """
    return list_id_and_name_member


def get_list_urgent_prj():
    """ return list urgent project """
    return list_urgent


def get_latest_comment(content, performer, comment_date, register_date):
    """ get formatted latest comment """
    today = date.today()
    if not content:
        if register_date == today:
            # if register today
            return "New Issue"
        # issue has no comment
        return "No Comment"
    result = str(content) + "   (" + str(performer) + "  " + str(comment_date) + ")"
    return result


def convert_dict_to_arr(dict_data):
    """ convert dict data to arr """
    out_data = [["", ""]]
    for key, value in dict_data.items():
        out_data.append([key, value])

    # Sort by team name A -> Z
    out_data.sort(key=lambda item: item[0])
    return out_data


def issue_analysis():
    """ analysis total issue report """
    long_pending_main = []
    long_pending_sub_mr = []

    pending_day = {}
    main_folder_issue = {}
    sub_mr_folder_issue = {}
    pending_issue_amount_5 = {}
    pending_issue_amount_7 = {}
    pending_sub_mr_by_team = {}

    global list_urgent
    list_urgent = get_urgent_project()
    list_flagship = get_flagship_prject()
    urgent_project = [item.upper() for item in list_urgent]
    # flagship_project = [item.upper() for item in list_flagship]
    flagship_project = list_flagship    # check flagship no need uppercase

    urgent_by_prj = {urgent_prj.lower(): [] for urgent_prj in urgent_project}
    flagship_by_prj = {flagship_prj.lower(): [] for flagship_prj in flagship_project}

    global flagship_issue_prj_count
    global urgent_issue_prj_count
    global owner_issue_urgent
    global owner_issue_pending
    global dfExcel
    global list_team_of_tg
    global member_id_by_team
    global list_team
    global list_tg
    global list_id_and_name_member

    list_tg, list_team_of_tg, list_team, numOfMember, \
    member_name_by_team, member_id_by_team, all_member_id_part, list_id_and_name_member = get_team_info()
    num_issue_of_member = dict.fromkeys(all_member_id_part, 0)
    issue_today_summary = {key: [0]*3 for key in list_team}    # ['today new', 'today resolve', 'today reject']

    for team in list_team:
        pending_day[team] = [0, 0, 0]
        pending_issue_amount_5[team] = 0
        pending_issue_amount_7[team] = 0
        pending_sub_mr_by_team[team] = 0
        #todo 04 June 2020  main_folder_issue[team] = [open, close, resolved]
        main_folder_issue[team] = [0, 0, 0]  # [open, close, resolved]
        sub_mr_folder_issue[team] = [0, 0, 0]  # [open, close, resolved]

    dfExcel, length_frame = read_excel_file(LIST_ISSUE, "DEFECT")

    pending_time_5 = 0	# Change 5 > 0 due to Request show all pending issue
    pending_time_7 = 0  # Change 7 > 0 due to Request show all pending issue
    TeamDetail = DataTabTeamDetail(list_team_of_tg)
    for i in range(0, length_frame):
        try:
            id_member = dfExcel['Manager ID'][i]
        except:
            continue
        team = team_member(str(id_member), member_id_by_team)
        tg = team_member(str(team), list_team_of_tg)

        if team is not None:
            # count issue open
            issue = [dfExcel['Case Code'][i], dfExcel['Dev. Mdl. Name/Item Name'][i],
                     dfExcel['PPT'][i], dfExcel['Title'][i],
                     dfExcel['Manager ID'][i], team, tg]
            # dfExcel['Pjt. Name'][i], dfExcel['Priority'][i],
            issue[0] = Wiki.makeLinkPLM(issue[0])
            issue[4] = Wiki.makeLinkChat(issue[4])

            checkUrgent = is_urgent_issue(dfExcel['Dev. Mdl. Name/Item Name'][i], urgent_project).lower()
            checkFlagship = is_urgent_issue(dfExcel['Dev. Mdl. Name/Item Name'][i], flagship_project).lower()

            if dfExcel['PPT'][i] != '-':
                num_issue_of_member[id_member] += 1
                is_folder_issue = folder_issue(dfExcel, i)

                # - START get comment Of issue (data table Team Detail)
                tg_name_key = tg.replace(' ', '_')
                name_model = dfExcel['Dev. Mdl. Name/Item Name'][i]
                comment_content = dfExcel['Detailed Information'][i]
                comment_performer = dfExcel['Performer'][i]
                comment_date = dfExcel['Reg. Date'][i]
                register_date = datetime.strptime(str(dfExcel['Registered Date'][i])[:10], '%Y-%m-%d').date()
                team_issue_detail = issue + [get_latest_comment(comment_content, comment_performer, comment_date, register_date)]
                # - END: data table Team Detail

                if register_date == date.today():
                    # count new issue register today
                    issue_today_summary[team][0] += 1

                # get issue detail by project add Column Daily comment
                urgent_issue = issue + [get_latest_comment(comment_content, comment_performer, comment_date, register_date)]
                # START: count Flagship issue include MR/Sub folder
                if checkFlagship != '' and is_folder_issue == 'Main':
                    if checkFlagship in flagship_issue_prj_count:
                        if team in flagship_issue_prj_count[checkFlagship]:
                            flagship_issue_prj_count[checkFlagship][team][0] += 1  # count flagship open issue
                        else:
                            flagship_issue_prj_count[checkFlagship][team] = [1, 0]  # flagship[Open, today_resolved]
                    else:
                        flagship_issue_prj_count[checkFlagship] = {}
                        flagship_issue_prj_count[checkFlagship][team] = [1, 0]

                    flagship_by_prj[checkFlagship].append(urgent_issue)

                if is_folder_issue == 'Main':
                    main_folder_issue[team][0] += 1  # count open issue Main Folder
                    if checkUrgent != '':
                        if checkUrgent in urgent_issue_prj_count:
                            if team in urgent_issue_prj_count[checkUrgent]:
                                urgent_issue_prj_count[checkUrgent][team][0] += 1  # count urgent open issue
                            else:
                                urgent_issue_prj_count[checkUrgent][team] = [1, 0]  # urgent[Open, today resolved]
                        else:
                            urgent_issue_prj_count[checkUrgent] = {}
                            urgent_issue_prj_count[checkUrgent][team] = [1, 0]

                        urgent_by_prj[checkUrgent].append(urgent_issue)

                        if id_member not in owner_issue_urgent:
                            owner_issue_urgent.append(id_member)

                    # count issue long pending
                    pday = [int(s) for s in re.findall('\\d+', dfExcel['PPT'][i])]
                    for idx in range(3):
                        pending_day[team][idx] += pday[idx]
                    if is_pending_issue(dfExcel['PPT'][i], pending_time_5):
                        pending_issue_amount_5[team] += 1
                        long_pending_main.append(team_issue_detail)
                        if id_member not in owner_issue_pending:
                            owner_issue_pending.append(id_member)
                    if is_pending_issue(dfExcel['PPT'][i], pending_time_7):
                        pending_issue_amount_7[team] += 1

                    # - START: data Tab Team Detail chart pie Main folder
                    TeamDetail.create_data_team_main_sub(team, name_model, 'Main')
                    TeamDetail.dict_info_issues_open_by_team[team]['main_folder'].append(team_issue_detail)
                    TeamDetail.create_data_tg_main_sub(tg_name_key, name_model, 'Main')
                    TeamDetail.dict_info_issues_open_by_tg[tg_name_key]['main_folder'].append(team_issue_detail)
                else:
                    # count OPEN issue MR and Sub folder
                    sub_mr_folder_issue[team][0] += 1
                    if is_pending_issue(dfExcel['PPT'][i], pending_time_7):
                        long_pending_sub_mr.append(team_issue_detail)
                        pending_sub_mr_by_team[team] += 1

                    # - START: data Tab Team Detail chart pie Main folder
                    TeamDetail.create_data_team_main_sub(team, name_model, 'Sub')
                    TeamDetail.dict_info_issues_open_by_team[team]['sub_folder'].append(team_issue_detail)
                    TeamDetail.create_data_tg_main_sub(tg_name_key, name_model, 'Sub')
                    TeamDetail.dict_info_issues_open_by_tg[tg_name_key]['sub_folder'].append(team_issue_detail)
            else:
                # count issue resolve today
                resolve_time = str(dfExcel['Resolve Date'][i])[:10]
                try:
                    resolve_date = datetime.strptime(resolve_time, '%Y-%m-%d').date()
                    if resolve_date == date.today():
                        # number issue of member include resolved today
                        num_issue_of_member[id_member] += 1
                        issue_today_summary[team][1] += 1   # count resolve today of team

                        # count urgent issue today resolved
                        if checkUrgent != '':
                            if checkUrgent in urgent_issue_prj_count:
                                if team in urgent_issue_prj_count[checkUrgent]:
                                    urgent_issue_prj_count[checkUrgent][team][1] += 1  # count urgent open issue
                                else:
                                    urgent_issue_prj_count[checkUrgent][team] = [0, 1]  # urgent [Open, today resolved]
                            else:
                                urgent_issue_prj_count[checkUrgent] = {}
                                urgent_issue_prj_count[checkUrgent][team] = [0, 1]

                        # START: count Flagship today resolved
                        if checkFlagship != '':
                            if checkFlagship in flagship_issue_prj_count:
                                if team in flagship_issue_prj_count[checkFlagship]:
                                    flagship_issue_prj_count[checkFlagship][team][1] += 1  # count flagship resolved
                                else:
                                    flagship_issue_prj_count[checkFlagship][team] = [0, 1]  # flagship today resolved
                            else:
                                flagship_issue_prj_count[checkFlagship] = {}
                                flagship_issue_prj_count[checkFlagship][team] = [0, 1]
                except:
                    # print("Error get resolve time case code: ", dfExcel['Case Code'][i])
                    pass

                # count sub and MR issue resolved / close
                if str(dfExcel['Resloved Confirm Date'][i]) == 'NaT':
                    if folder_issue(dfExcel, i) == 'Main':
                        main_folder_issue[team][2] += 1  # count issue RESOLVED not close in Main Folder
                    else:
                        sub_mr_folder_issue[team][2] += 1
                else:
                    # count issue CLOSE in Main Folder
                    if folder_issue(dfExcel, i) == 'Main':
                        main_folder_issue[team][1] += 1
                    else:
                        sub_mr_folder_issue[team][1] += 1

            # count num issue reject today
            try:
                num_of_reject = int(dfExcel['# of Reject'][i])
                if num_of_reject > 0:
                    reject_time = str(dfExcel['Reject Date'][i])[:10]
                    reject_date = datetime.strptime(reject_time, '%Y-%m-%d').date()
                    if reject_date == date.today():
                        issue_today_summary[team][2] += 1
            except:
                pass

    data_chart_today_new = []
    for tg, team_of_tg in list_team_of_tg.items():
        list_value = [issue_today_summary.get(key) for key in team_of_tg]
        dict_data_tg = {team: value for team, value in zip(team_of_tg, list_value)}
        dict_data_tg['tg_name'] = tg
        data_chart_today_new.append(dict_data_tg)

    long_pending_sub_mr[1:] = sorted(long_pending_sub_mr[1:], key=lambda x: -ppt_min(x))
    # subprocess.call("taskkill /IM excel.exe")

    pending_result = {}

    for team in list_team:
        pending_result[team] = pending_day[team][0] + int(pending_day[team][1] / 24) + int(
            pending_day[team][2] / (24 * 60))

    # Create data chart long pending Main/Sub in PLM Issue TAB
    data_pending_main_sub = {}
    data_pending_main_sub['main_5days'] = convert_dict_to_arr(pending_issue_amount_5)
    data_pending_main_sub['sub_7days'] = convert_dict_to_arr(pending_sub_mr_by_team)
    # create data table long pending Main 5days / Sub 7days
    data_table_long_pending = {}
    data_table_long_pending['main_5days'] = long_pending_main
    data_table_long_pending['sub_7days'] = long_pending_sub_mr

    pending_issue_amount_5['type_data'] = "fiveDays"
    pending_issue_amount_7['type_data'] = "sevenDays"
    num_pending_57days = [pending_issue_amount_5, pending_issue_amount_7]

    main_plm_open = summary_analysis(main_folder_issue)
    main_summary_bar = summary_analysis_active(main_folder_issue, 'main')
    sub_summary_bar = summary_analysis_active(sub_mr_folder_issue, 'sub')
    summary_data = main_summary_bar + sub_summary_bar

    # add new item detail all urgent projects
    all_urgent_prj = []
    for key, values in urgent_by_prj.items():
        all_urgent_prj += values
    urgent_by_prj['detail_all_urgent'] = all_urgent_prj

    all_flagship_prj = []
    for key, values in flagship_by_prj.items():
        all_flagship_prj += values
    flagship_by_prj['detail_all_flagship'] = all_flagship_prj

    data_bar_chart_of_team_detail = TeamDetail.get_data_bar_chart()
    dict_info_issues_open_by_team = TeamDetail.get_data_table()

    return main_plm_open, summary_data, data_table_long_pending, \
           num_pending_57days, data_pending_main_sub, \
           urgent_by_prj, flagship_by_prj, list_team, all_member_id_part, member_id_by_team, \
           data_bar_chart_of_team_detail, dict_info_issues_open_by_team, num_issue_of_member, data_chart_today_new
