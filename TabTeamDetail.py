class DataTabTeamDetail:
    dict_issue_open_by_team_main = {}
    dict_issue_open_by_team_sub_mr = {}
    dict_info_issues_open_by_team = {}

    dict_issue_open_by_tg_main = {}
    dict_issue_open_by_tg_sub_mr = {}
    dict_info_issues_open_by_tg = {}

    dict_team_of_tg = {}
    list_team_name = []
    list_tg_name = []

    def __init__(self, list_team_by_tg):
        self.dict_team_of_tg = list_team_by_tg
        key_sub_main = ['main_folder', 'sub_folder']
        for key, value in list_team_by_tg.items():
            tg_key_name = key.replace(' ', '_')
            self.list_tg_name.append(tg_key_name)
            self.list_team_name += value
            self.dict_issue_open_by_tg_main[tg_key_name] = {}
            self.dict_issue_open_by_tg_sub_mr[tg_key_name] = {}
            self.dict_info_issues_open_by_tg[tg_key_name] = {key_sub: [] for key_sub in key_sub_main}
            for team in value:
                self.dict_issue_open_by_team_main[team] = {}
                self.dict_issue_open_by_team_sub_mr[team] = {}
                self.dict_info_issues_open_by_team[team] = {key_name: [] for key_name in key_sub_main}

    def create_data_team_main_sub(self, team, name_model, type_issue):
        """ create data issue main folder """
        if name_model.startswith('Galaxy'):
            name_model = name_model.split()[0]
        elif 'SM-' in name_model:
            pos = name_model.index('SM-')
            name_model = name_model[pos + 3:]
            name_model = name_model.split('_')[0]
        else:
            name_model = 'Others Project'

        if type_issue == 'Main':
            try:
                self.dict_issue_open_by_team_main[team][name_model] = \
                    self.dict_issue_open_by_team_main[team][name_model] + 1
            except KeyError:
                self.dict_issue_open_by_team_main[team][name_model] = 1
        else:
            try:
                self.dict_issue_open_by_team_sub_mr[team][name_model] = \
                    self.dict_issue_open_by_team_sub_mr[team][name_model] + 1
            except KeyError:
                self.dict_issue_open_by_team_sub_mr[team][name_model] = 1

    def create_data_tg_main_sub(self, tg_name, name_model, type_issue):
        """ create data issue main folder """
        if name_model.startswith('Galaxy'):
            name_model = name_model.split()[0]
        elif 'SM-' in name_model:
            pos = name_model.index('SM-')
            name_model = name_model[pos + 3:]
            name_model = name_model.split('_')[0]
        else:
            name_model = 'Others Project'

        if type_issue == 'Main':
            try:
                self.dict_issue_open_by_tg_main[tg_name][name_model] = \
                    self.dict_issue_open_by_tg_main[tg_name][name_model] + 1
            except KeyError:
                self.dict_issue_open_by_tg_main[tg_name][name_model] = 1
        else:
            try:
                self.dict_issue_open_by_tg_sub_mr[tg_name][name_model] = \
                    self.dict_issue_open_by_tg_sub_mr[tg_name][name_model] + 1
            except KeyError:
                self.dict_issue_open_by_tg_sub_mr[tg_name][name_model] = 1

    def export_data_bar_chart(self, dict_issue_main, dict_issue_sub):
        """ generate data for pie chart main_sub """
        result_data = {}
        other_project = 'Others Project'

        for key, value in dict_issue_main.items():
            arr_value_main = [[k, v] for k, v in value.items()]
            arr_value_main.sort(key=lambda item: item[1], reverse=True)
            if len(arr_value_main) > 10:
                sum_prj_less_issue = sum(item[1] for item in arr_value_main[10:])
                if other_project in value:
                    num_other_prj = value[other_project]
                    pos_other_prj = arr_value_main.index([other_project, num_other_prj])
                    arr_value_main[pos_other_prj] = [other_project, num_other_prj + sum_prj_less_issue]
                elif sum_prj_less_issue > 0:
                    arr_value_main.append([other_project, sum_prj_less_issue])
                temp_data = {key: {'main_folder': [['', 'Number issue']] + arr_value_main[:10]}}
            else:
                temp_data = {key: {'main_folder': [['', 'Number issue']] + arr_value_main}}
            result_data.update(temp_data)

        for key, value in dict_issue_sub.items():
            arr_value_sub = [[k, v] for k, v in value.items()]
            arr_value_sub.sort(key=lambda item: item[1], reverse=True)
            if len(arr_value_sub) > 10:
                sum_prj_less_issue = sum(item[1] for item in arr_value_sub[10:])
                if other_project in value:
                    num_other_prj = value[other_project]
                    pos_other_prj = arr_value_sub.index([other_project, num_other_prj])
                    arr_value_sub[pos_other_prj] = [other_project, num_other_prj + sum_prj_less_issue]
                elif sum_prj_less_issue > 0:
                    arr_value_sub.append([other_project, sum_prj_less_issue])
                temp_data = {'sub_folder': [['', 'Number issue']] + arr_value_sub[:10]}
            else:
                temp_data = {'sub_folder': [['', 'Number issue']] + arr_value_sub}
            if key in result_data:
                result_data[key].update(temp_data)
            else:
                result_data.update({key: temp_data})
            result_data.update(temp_data)
        return result_data

    def get_data_bar_chart(self):
        """ return data include tg and team """
        result_team = self.export_data_bar_chart(self.dict_issue_open_by_team_main, self.dict_issue_open_by_team_sub_mr)
        tg_data = self.export_data_bar_chart(self.dict_issue_open_by_tg_main, self.dict_issue_open_by_tg_sub_mr)
        result_team.update(tg_data)
        return result_team

    def get_data_table(self):
        """ return data table by team detail """
        self.dict_info_issues_open_by_team.update(self.dict_info_issues_open_by_tg)
        return self.dict_info_issues_open_by_team
