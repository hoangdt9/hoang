import traceback
import WikiSubmit as Wiki
import ReadDataFromExcel as DataEXL
import pandas as pd
import sys


def getIssueCount(prj):
    return DataEXL.urgent_count_prj(prj.lower())


def getFlagshipCount(prj):
    return DataEXL.flagship_count_prj(prj.lower())


def get_list_project(typePrj='Urgent'):
    """ get list project Urgent / Flagship """
    list_prj = []
    try:
        res = Wiki.getPageContent(Wiki.pageUrgentPrjTitle, Wiki.space_key)
        table = pd.read_html(res.content)
        if typePrj == 'Urgent':
            table = table[0]
        else:
            table = table[1]    # create flagship table
        for index, row in table[1:].iterrows():
            list_prj.append(row[1])

    except Exception as e:
        print(e)
        traceback.print_exc()
        print("Cannot get page %s content" % Wiki.pageUrgentPrjTitle)
        sys.exit(1)
    return list_prj


def getUrgentTableCode(typeTable='Urgent'):
    base_table = open('v2_template/table.html','r').read()
    prj_has_issue = []
    row_html = '''<tr class="{}">
                    <td class="prj-name">{}</td>
                    <td class="model-name">{}</td>
                    <td>{}</td>
                    <td>{}</td>
                    <td class="num-issue">{}</td>
                    <td>{}</td>
                </tr>'''
    try:
        res = Wiki.getPageContent(Wiki.pageUrgentPrjTitle, Wiki.space_key)
        table = pd.read_html(res.content)
        if typeTable == 'Urgent':
            table = table[0]
            base_table = base_table.replace('<!--table_id_name-->', 'table-urgent')
            base_table = base_table.replace('<!--table_name-->', "Urgent Projects")
        else:
            table = table[1]    # create flagship table
            base_table = base_table.replace('<!--table_id_name-->', 'table-flagship')
            base_table = base_table.replace('<!--table_name-->', "Flagship Models")
        total_issues = 0
        total_resolved = 0
        for index, row in table[1:].iterrows():
            prj = row[0]

            model = row[1]

            pvr = row[2]
            if str(pvr).lower() == 'nan':
                pvr = '-'

            pra = row[3]
            if str(pra).lower() == 'nan':
                pra = '-'

            position = base_table.find('<!--insert_row-->')
            if typeTable == 'Urgent':
                issues, today_resolved = getIssueCount(model)
            else:
                issues, today_resolved = getFlagshipCount(model)

            if issues != 0:
                total_issues += issues
                prj_type = 'table-danger'
                prj_has_issue.append(model.upper())
            else:
                prj_type = 'table-primary'
            total_resolved += today_resolved

            row = row_html.format(prj_type, prj, model, pvr, pra, issues, today_resolved)
            base_table = base_table[:position] + row + base_table[position:]

        latest_row = row_html.replace('class="prj-name"', '')
        row_total = latest_row.format('table-urgent-latest', 'Total Issues', '', '', '', total_issues, total_resolved)
        base_table = base_table.replace('<!--insert_row-->', row_total)

    except Exception as e:
        print(e)
        traceback.print_exc()
        print("Cannot get page %s content" % Wiki.pageUrgentPrjTitle)

    return base_table


# Create table summary flagship mobihub
def createTableFlagshipMobihub(flagshipMobihub):
    base_table = open('v2_template/table.html', 'r').read()
    row_html = '''<tr class="{}">
                        <td class="prj-name">{}</td>
                        <td class="model-name">{}</td>
                        <td>{}</td>
                        <td>{}</td>
                        <td class="num-issue">{}</td>
                        <td>{}</td>
                    </tr>'''
    base_table = base_table.replace('<!--table_id_name-->', 'table-flagship-mobihub')
    base_table = base_table.replace('<!--table_name-->', "Flagship Models")

    model = '-'
    prj_type = 'table-primary'
    resolved = 0
    total_issues = 0
    for prj_name, num_issue in flagshipMobihub.items():
        position = base_table.find('<!--insert_row-->')
        row = row_html.format(prj_type, prj_name, model, '-', '-', num_issue, resolved)
        base_table = base_table[:position] + row + base_table[position:]
        total_issues += num_issue

    latest_row = row_html.replace('class="prj-name"', '')
    row_total = latest_row.format('table-urgent-latest', 'Total Issues', '', '', '', total_issues, resolved)
    base_table = base_table.replace('<!--insert_row-->', row_total)

    return base_table

