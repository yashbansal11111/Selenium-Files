import threading
import queue
import time as t
from selenium import webdriver
from selenium.common import exceptions
import pandas as pd
from selenium.webdriver.common.by import By
import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

q = queue.Queue()

no_of_tabs = input("Enter the no. of tabs you want to open at a time: ")
attempts_case_1 = input("Enter the no. of attempts you want the script to perform for finding web element: ")

all_sheets_dict = pd.read_excel('response_sheet.xlsx', sheet_name=None)
reschedule_sheet = pd.read_excel("reschedule_sheet_small.xlsx")
no_of_rows_range = all_sheets_dict['Sheet1'].index
len_rows = len(no_of_rows_range)
initial = len_rows - int(no_of_tabs) - 1

regards_list = ["Thank you for your time and consideration; our HR staff will contact you for a follow-up interview.",
                "Thank you for your time and consideration; our Human Resources staff will contact you to schedule a follow-up interview.",
                "Thank you for your time and consideration; our Human Resources staff will contact you for a follow-up interview.",
                "Thank you for your time and consideration.",
                "Thank you for your valuable time and consideration.",
                "Thank you for your time."
                ]
reschedule_question_list = ['When can we call you again?',
                            'At what time you will be available for discussion?'
                            ]
gibberish_question_list = ["Sorry i didn't understand, please repeat",
                           "can you say it one more time?",
                           "I didn't get you",
                           ]
pass_question_list = ['Hello John, Very Good afternoon.',
                      'Hello John, Very Good morning.',
                      'Hello John, Very Good evening.',
                      'Hello John, Good morning.',
                      'Hello John, Good afternoon.',
                      'Hello John, Good evening.',
                      'I am calling from Dell Technologies. Is it good moment to talk?',
                      'I am phoning from Dell Technologies. Is it right time to talk?',
                      "We're looking for a Principal Software Engineer-IT to join our team based in Hyderabad. Are you interested?",
                      "We are searching for a Principal Software Engineer-IT to join our team in Hyderabad. Are you interested?",
                      "Principal Software Engineer-IT is a position that we are looking to fill for Hyderabad location. Are you interested?",
                      "How many years of python experience do you have?",
                      "How comfortable you are with python?",
                      "How much you are making in current job?",
                      "how much you are making right now?",
                      "What is your renumeration currently?",
                      "how much you get paid right now?",
                      "How much do you expect to be paid?",
                      "What are your salary goals?",
                      "What are your salary targets?",
                      "What's your salary targets?",
                      "When will you be available for the next round of interviews?",
                      "When will the next round of interviews be convenient for you?",
                      "What time will you be available for the next round of interviews?"
                      ]
                    
driver = webdriver.Firefox()

def webdriverhandler(driver, all_sheets_dict, i, kwargs):
    sent_response_list = kwargs['sent_response_list']
    chat_result_df = kwargs['chat_result_df']
    inner_flag = kwargs['inner_flag']
    outer_flag = kwargs['outer_flag']
    sheet_no = kwargs['sheet_no']
    chat_history_list = kwargs['chat_history_list']
    time_logs_list = kwargs['time_logs_list']
    chatbot_element_counter = kwargs['chatbot_element_counter']
    reschedule_index_increment = kwargs['reschedule_index_increment']
    conv_no = kwargs['conv_no']
    main_writer = kwargs['main_writer']
    reschedule_writer = kwargs['reschedule_writer']
    for k,v in all_sheets_dict.items():
        k = pd.DataFrame(v)
        try:
            reply_field = driver.find_element(by=By.XPATH, value='/html/body/div/div/div/div[2]/div/input')
            reply_field.send_keys(k.iloc[i,1])
            reply_button = driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div/div[2]/div/button')
            reply_button.click()
            start_time = datetime.datetime.now()
            sent_response_list.append(k.iloc[i, 1])
            for attempt in range(1, int(attempts_case_1) + 1):
                print(f'Attempt No: {attempt}')
                try:
                    open_text_component_element_xpath = f"//div[@id='conversationSection']/div[@class='open-text-component'][{chatbot_element_counter}]/div[@class='ant-row']/div[@class='message-control bot-control']/div[@class='msg-box default-control']"
                    element_obj = WebDriverWait(driver, 4).until(
                                EC.presence_of_element_located((By.XPATH, open_text_component_element_xpath)
                                                                )
                                                                )
                    element = element_obj.text
                    if element:
                        time_log = datetime.datetime.now() - start_time
                        time_logs_list.append(time_log)            

                    first_ques = driver.find_element(by=By.XPATH, value="/html/body/div/div/div/div[1]/div[2]/div[1]/div/div/span").text

                    if element in gibberish_question_list:
                        k.iloc[i, 2] = 'Failed'
                        inner_flag = True
                        outer_flag = True
                    elif element == first_ques or element in chat_history_list:
                        k.iloc[i, 2] = 'Repeated'
                        inner_flag = True
                        outer_flag = True
                    elif element in pass_question_list:
                        k.iloc[i, 2] = 'Passed'
                    elif element in reschedule_question_list:
                        k.iloc[i, 2] = 'Rescheduled'
                        for j in range(len(reschedule_sheet.index)):
                            reply_field = driver.find_element(by=By.XPATH, value='/html/body/div/div/div/div[2]/div/input')
                            reply_field.send_keys(reschedule_sheet.iloc[j+reschedule_index_increment,1])
                            reply_button = driver.find_element(by=By.XPATH, value='/html/body/div[1]/div/div/div[2]/div/button')
                            reply_button.click()
                            start_time = datetime.datetime.now()
                            sent_response_list.append(reschedule_sheet.iloc[j+reschedule_index_increment,1])
                            for reschedule_attempt in range(1, int(attempts_case_1) + 1):
                                print(f'Reschedule Attempt No: {reschedule_attempt}')
                                try:
                                    message_component_element_xpath = "//div[@id='conversationSection']/div[@class='message-component'][2]/div[@class='ant-row']/div[@class='message-control bot-control']/div[@class='msg-box default-control']"
                                    reschedule_element_obj = WebDriverWait(driver, 4).until(
                                                            EC.presence_of_element_located((By.XPATH, message_component_element_xpath)
                                                                                            )
                                                                                            )
                                    reschedule_element = reschedule_element_obj.text
                                    if reschedule_element:
                                        time_log = datetime.datetime.now() - start_time
                                        time_logs_list.append(time_log)
                                    if reschedule_element in regards_list:
                                        reschedule_sheet.iloc[j+reschedule_index_increment,2] = 'Passed'
                                    reschedule_sheet.to_excel(reschedule_writer, index=False)
                                    reschedule_writer.save()
                                    break
                                except TimeoutException:
                                    continue
                            else:
                                time_log = datetime.datetime.now() - start_time
                                time_logs_list.append(time_log)
                                reschedule_sheet.iloc[j+reschedule_index_increment, 2] = 'Server Timeout'
                                reschedule_sheet.to_excel(reschedule_writer, index=False)
                                reschedule_writer.save()
                            reschedule_index_increment = reschedule_index_increment + 1
                            inner_flag = True
                            break
                    else:
                        k.iloc[i, 2] = 'Paragraph Repeat'
                        inner_flag = True
                        outer_flag = True
                    t.sleep(0.5)
                    k.to_excel(main_writer, sheet_name=f'status_sheet{sheet_no}', index=False)
                    sheet_no = sheet_no + 1
                    main_writer.save()
                    if inner_flag:
                        inner_flag = False
                        break
                    chat_history_list.append(element)
                    chatbot_element_counter = chatbot_element_counter + 1
                    break
                except TimeoutException:
                    start_time_for_message_component_element = datetime.datetime.now()
                    try:
                        message_component_element_xpath = "//div[@id='conversationSection']/div[@class='message-component'][2]/div[@class='ant-row']/div[@class='message-control bot-control']/div[@class='msg-box default-control']"
                        message_component_element_obj = WebDriverWait(driver, 4).until(
                                                        EC.presence_of_element_located((By.XPATH, message_component_element_xpath)
                                                                                    )
                                                                                    )
                        message_component_element = message_component_element_obj.text
                        if message_component_element:
                            time_log = datetime.datetime.now() - start_time_for_message_component_element
                            time_logs_list.append(time_log)
                        if message_component_element in regards_list:
                            k.iloc[i, 2] = 'Passed'
                            k.to_excel(main_writer, sheet_name=f'status_sheet{sheet_no}', index=False)
                            main_writer.save()
                            break
                    except TimeoutException:
                        print(f'Attempt {attempt} hit the exception.')
                        continue
                    continue
            else:
                time_log = datetime.datetime.now() - start_time
                time_logs_list.append(time_log)  
                k.iloc[i, 2] = 'Server Timeout'
                k.to_excel(main_writer, sheet_name=f'status_sheet{sheet_no}', index=False)
                main_writer.save()
                break
        except exceptions.NoSuchElementException:
            driver.implicitly_wait(2)
        if outer_flag:
            outer_flag = False
            break

    # chat_results.csv
    chats = driver.find_elements(By.XPATH, "//span[contains(@class,'msg-text')]")
    t.sleep(0.5)
    datetime_stamp = datetime.datetime.now()
    timestamp = "Timestamp: %s-%s-%s %s:%s:%s" % (datetime_stamp.year,
                                                  datetime_stamp.month,
                                                  datetime_stamp.day,
                                                  datetime_stamp.hour, 
                                                  datetime_stamp.minute, 
                                                  datetime_stamp.second 
                                                  )
    chatbot_response_list = [conv_no,timestamp]
    user_response_list = ['', '', '']
    for item in chats:
        chat_in_textformat = item.text
        if chat_in_textformat not in sent_response_list:
            chatbot_response_list.append(chat_in_textformat)
        else:
            user_response_list.append(chat_in_textformat)
    chatbot_response_list.append('')
    user_response_list.append('')
    user_response_list.append('')
    chat_result_df['Questions'] = pd.Series(chatbot_response_list)
    chat_result_df['Replies'] = pd.Series(user_response_list)
    chat_result_df['Timelog'] = pd.Series(time_logs_list)
    chat_result_df.to_csv('chat_results.csv', mode='a', index=False)
    conv_no = conv_no + 1
    driver.implicitly_wait(2)
    return True
    # driver.refresh()
    # else:
    #     print('There seems to be an error in TimeStamp')

# Looping through all the sheets by rows.
def rootle_automation(status_file, reschedule_file, start_row_number=0, end_row_number=len_rows, tab_no=1):
    python_kwrds = dict()
    reschedule_index_increment = 0
    conv_no = 1
    main_writer = pd.ExcelWriter(status_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') # pylint: disable=abstract-class-instantiated
    reschedule_writer = pd.ExcelWriter(reschedule_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') # pylint: disable=abstract-class-instantiated
    python_kwrds['reschedule_index_increment'] = reschedule_index_increment
    python_kwrds['conv_no'] = conv_no
    python_kwrds['main_writer'] = main_writer
    python_kwrds['reschedule_writer'] = reschedule_writer
    # tab_no = 1
    # for tab in range(1, int(no_of_tabs) + 1):
    driver.maximize_window()
    driver.execute_script(f"window.open('about:blank','Tab{tab_no}');")
    driver.switch_to.window(f"Tab{tab_no}")
    driver.get('https://chat.rootle.ai/')
    driver.implicitly_wait(3)
    # start_row_number = start_row_number+tab_no
    for i in range(start_row_number, end_row_number):
        print('Iteration No.: ',i)
        sent_response_list = []
        chat_result_df = pd.DataFrame(columns=['Questions',
                                            'Replies',
                                            'Timelog'
                                            ]
                                    )
        inner_flag = False
        outer_flag = False
        sheet_no = 1
        chat_history_list = []
        time_logs_list = ['', '', '']
        chatbot_element_counter = 2
        python_kwrds['sent_response_list'] = sent_response_list
        python_kwrds['chat_result_df'] = chat_result_df
        python_kwrds['inner_flag'] = inner_flag
        python_kwrds['outer_flag'] = outer_flag
        python_kwrds['sheet_no'] = sheet_no
        python_kwrds['chat_history_list'] = chat_history_list
        python_kwrds['time_logs_list'] = time_logs_list
        python_kwrds['chatbot_element_counter'] = chatbot_element_counter
        response = webdriverhandler(driver, all_sheets_dict, i, python_kwrds)
        # start_row_number +=1
        if response:
            break
    # tab_no += 1

            # # NEEDED...
            # # # Check greeting message on the basis of time-stamp.
            # # time_stamp_flag = False
            # # # greeting_msg = driver.find_element_by_class_name('msg-text').text
            # # greeting_msg = driver.find_element(by=By.CLASS_NAME, value='msg-text').text
            # # am_or_pm = driver.find_element_by_class_name('time-stamp').text.split(' ')
            # # # am_or_pm = driver.find_element(by=By.CLASS_NAME, value='time-stamp').text.split(' ')
            # # if am_or_pm[1] == 'pm':
            # #     time = int(am_or_pm[0].split(':')[0]) + 12
            # #     if time>=11 and time<=17 or time == 24:
            # #         if greeting_msg in ['Hello John, Very Good afternoon.', 'Hello John, Good afternoon.']:
            # #             time_stamp_flag = True
            # #     if time>=18 and time<=23:
            # #         if greeting_msg in ['Hello John, Very Good evening.', 'Hello John, Good evening.']:
            # #             time_stamp_flag=True
            # # elif am_or_pm[1] == 'am':
            # #     if greeting_msg in ['Hello John, Very Good morning.', 'Hello John, Good morning.']:
            # #         time_stamp_flag = True
            # if time_stamp_flag:
            
# rootle_automation("response_with_status_temp.xlsx", "reschedule_sheet_small.xlsx", initial, len_rows)

# thread_list = list()

def worker():
    while True:
        item = q.get()
        # print(f'Working on {item}')
        # print(f'Finished {item}')
        q.task_done()

# Turn-on the worker thread.
threading.Thread(target=worker, daemon=True).start()

for tmp in range(1, int(no_of_tabs) + 1):
    initial += 1
    q.put(rootle_automation("response_with_status_temp.xlsx", "reschedule_sheet_small.xlsx", initial, len_rows, tmp))
    # th = threading.Thread(name='Test {}'.format(tmp), target=rootle_automation("response_with_status_temp.xlsx", "reschedule_sheet_small.xlsx", initial, len_rows, tmp))
    # th.start()
    # # t.sleep(1)
    # print(th.name + 'started')
    # thread_list.append(th)


# for thread in thread_list:
#     thread.join()
q.join()
print('All work completed')




