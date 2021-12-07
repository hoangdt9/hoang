import WikiSubmit as Wiki
from datetime import datetime
import os


def is_update_success():
    """ check if run daily report successfully """
    file_success = os.getcwd() + "\\last_success.txt"
    if os.path.exists(file_success):
        return True
    else:
        return False


def check_to_update():
    """ program interval """
    time_update = Wiki.check_time_update()
    current_time = datetime.now()
    time_delta = current_time - time_update
    time_delta_min = time_delta.seconds // 60   # convert to minutes;
    print("daily report update : %d minute ago" % time_delta_min)
    #todo 04 June 2020 check time update to wiki
    time_delays = 35 #35 - time_delta_min

    # todo 04 June 2020 check time update to wiki
    if time_delta_min > 0:  #30
        print(" start run Daily Report Backup")
        os.system('python V2_Home.py')
        if is_update_success():
            print("run daily report successful!")
            # todo 04 June 2020 check time update to wiki
            time_delays = 35
        else:
            print("try again after 2 minutes")
            time_delays = 2

    print("start schedule check update after: %d minutes" % time_delays)
    return time_delays * 60


if __name__ == "__main__":
    time_alarm = check_to_update()
    os._exit(time_alarm)
