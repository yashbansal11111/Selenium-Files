import time as t
from selenium import webdriver
from selenium.common import exceptions
import pandas as pd
from selenium.webdriver.common.by import By
import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


browser = webdriver.Firefox()
browser.maximize_window()
browser.get('https://chat.rootle.ai/')
browser.implicitly_wait(3)

all_sheets_dict = pd.read_excel('response_sheet_small.xlsx', sheet_name=None)
reschedule_sheet = pd.read_excel("reschedule_sheet_small.xlsx")
no_of_rows_range = all_sheets_dict['Sheet1'].index
len_rows = len(no_of_rows_range)

#Check greeting message on the basis of time-stamp.
# time_stamp_flag = False
# greeting_msg = browser.find_element_by_class_name('msg-text').text
# am_or_pm = browser.find_element_by_class_name('time-stamp').text.split(' ')
# if am_or_pm[1] == 'pm':
#     time = int(am_or_pm[0].split(':')[0]) + 12
#     if time>=11 and time<=17 or time == 24:
#         if greeting_msg in ['Hello John, Very Good afternoon.', 'Hello John, Good afternoon.']:
#             time_stamp_flag = True
#     if time>=18 and time<=23:
#         if greeting_msg in ['Hello John, Very Good evening.', 'Hello John, Good evening.']:
#             time_stamp_flag=True
# elif am_or_pm[1] == 'am':
#     if greeting_msg in ['Hello John, Very Good morning.', 'Hello John, Good morning.']:
#         time_stamp_flag = True

regards_list = ["Thank you for your time and consideration; our HR staff will contact you for a follow-up interview.",
                "Thank you for your time and consideration; our Human Resources staff will contact you to schedule a follow-up interview.",
                "Thank you for your time and consideration; our Human Resources staff will contact you for a follow-up interview."
                ]
reschedule_question_list = ['When can we call you again?',
                            'At what time you will be available for discussion?'
                            ]
gibberish_question_list = ["Sorry i didn't understand, please repeat",
                           "can you say it one more time?",
                           "I didn't get you",
                           "Thank you for your time and consideration.",
                           "Thank you for your valuable time and consideration.",
                           "Thank you for your time."
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
                    "What time will you be available for the next round of interviews?",
                    "Thank you for your time and consideration; our HR staff will contact you for a follow-up interview.",
                    "Thank you for your time and consideration; our Human Resources staff will contact you to schedule a follow-up interview.",
                    "Thank you for your time and consideration; our Human Resources staff will contact you for a follow-up interview."
                    ]

#Looping through all the sheets by rows.
def automation_script(status_file, reschedule_file, start_row_number = 0, end_row_number = len_rows):
    sent_replies_list = []
    reschedule_index_increment = 0
    conv_no = 1
    writer = pd.ExcelWriter(status_file, engine = 'openpyxl', mode = 'a', if_sheet_exists = 'overlay') # pylint: disable=abstract-class-instantiated
    reschedule_writer = pd.ExcelWriter(reschedule_file, engine = 'openpyxl', mode = 'a', if_sheet_exists = 'overlay') # pylint: disable=abstract-class-instantiated
    # if time_stamp_flag:
    for i in range(start_row_number, end_row_number):
        print(i,'++++++++++')
        result_df = pd.DataFrame(columns=['Questions','Replies','Timelog'])
        flag = False
        tmp = 1
        chat_history_list = []
        time_logs_list = ['','','']
        chatbot_element_counter = 2
        for k,v in all_sheets_dict.items():
            k = pd.DataFrame(v)
            try:
                reply_field = browser.find_element_by_xpath('/html/body/div/div/div/div[2]/div/input')
                reply_field.send_keys(k.iloc[i,1])
                reply_button = browser.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/button')
                reply_button.click()
                t.sleep(0.5)
                start_time = datetime.datetime.now()
                sent_replies_list.append(k.iloc[i,1])
                try:
                    element_obj = WebDriverWait(browser, 7).until(
                                  EC.presence_of_element_located((By.XPATH, f"//div[@id='conversationSection']/div[@class='open-text-component'][{chatbot_element_counter}]/div[@class='ant-row']/div[@class='message-control bot-control']/div[@class='msg-box default-control']"))
                                  )
                    element = element_obj.text
                    if element:
                        time_log = datetime.datetime.now() - start_time
                        time_logs_list.append(time_log)                  

                    first_ques = browser.find_element_by_xpath("/html/body/div/div/div/div[1]/div[2]/div[1]/div/div/span").text

                    # all_element_list = browser.find_elements_by_class_name("msg-text")
                    # current_chat = all_element_list[-1].text
                    # for chat in all_element_list:
                    #     current_chat = chat.text

                    if element in gibberish_question_list:
                        k.iloc[i,2] = 'Failed'
                        flag = True
                    elif element == first_ques or element in chat_history_list:
                        t.sleep(0.5)
                        k.iloc[i,2] = 'Repeated'
                        flag=True
                    elif element in pass_question_list:
                        k.iloc[i,2] = 'Passed'
                    elif element in reschedule_question_list:
                        k.iloc[i,2] = 'Rescheduled'
                        for j in range(len(reschedule_sheet.index)):
                            reply_field = browser.find_element_by_xpath('/html/body/div/div/div/div[2]/div/input')
                            reply_field.send_keys(reschedule_sheet.iloc[j+reschedule_index_increment,1])
                            reply_button = browser.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/button')
                            reply_button.click()
                            # t.sleep(0.5)
                            start_time = datetime.datetime.now()
                            sent_replies_list.append(reschedule_sheet.iloc[j+reschedule_index_increment,1])

                            # all_element_list_for_rescheduling = browser.find_elements_by_class_name("msg-text")
                            # for chat in all_element_list_for_rescheduling:
                            #     current_chat_for_rescheduling = chat.text

                            try:
                                reschedule_element_obj = WebDriverWait(browser, 7).until(
                                            EC.presence_of_element_located((By.XPATH, f"//div[@id='conversationSection']/div[@class='message-component'][2]/div[@class='ant-row']/div[@class='message-control bot-control']/div[@class='msg-box default-control']"))
                                )
                                reschedule_element = reschedule_element_obj.text
                                if reschedule_element:
                                    time_log = datetime.datetime.now() - start_time
                                    time_logs_list.append(time_log)
                                if reschedule_element in regards_list:
                                    reschedule_sheet.iloc[j+reschedule_index_increment,2] = 'Passed'
                                t.sleep(0.5)
                                reschedule_sheet.to_excel(reschedule_writer, index=False)
                                reschedule_writer.save()
                            except TimeoutException as ex:
                                print('went to exception timeout for reschedule*******************************')
                                reschedule_sheet.iloc[j+reschedule_index_increment,2] = 'Server Timeout'
                                t.sleep(0.5)
                                reschedule_sheet.to_excel(reschedule_writer, index=False)
                                reschedule_writer.save()
                            reschedule_index_increment = reschedule_index_increment + 1
                            flag = True
                            break
                    else:
                        t.sleep(0.5)
                        k.iloc[i,2] = 'Paragraph Repeat'
                        flag = True
                    t.sleep(0.5)
                    k.to_excel(writer, sheet_name=f'status_sheet{tmp}', index=False)
                    tmp = tmp + 1
                    writer.save()
                    if flag:
                        flag = False
                        break
                    chat_history_list.append(element)
                    chatbot_element_counter = chatbot_element_counter + 1
                except TimeoutException:
                    try:
                        message_component_element_obj = WebDriverWait(browser, 7).until(
                                                        EC.presence_of_element_located((By.XPATH, "//div[@id='conversationSection']/div[@class='message-component'][2]/div[@class='ant-row']/div[@class='message-control bot-control']/div[@class='msg-box default-control']"))
                                                        )
                        message_component_element = message_component_element_obj.text
                        if message_component_element:
                            time_log = datetime.datetime.now() - start_time
                            time_logs_list.append(time_log)
                        if message_component_element in regards_list:
                            k.iloc[i,2] = 'Passed'
                            k.to_excel(writer, sheet_name=f'status_sheet{tmp}', index=False)
                            writer.save()
                            print(k)
                            break
                    except TimeoutException:
                        print('went to message component element exception timeout*******************************')
                        k.iloc[i,2] = 'Server Timeout'
                        t.sleep(0.5)
                        k.to_excel(writer, sheet_name=f'status_sheet{tmp}', index=False)
                        writer.save()
                        print(k)
                    print('went to exception timeout*******************************')
                    k.iloc[i,2] = 'Server Timeout'
                    t.sleep(0.5)
                    k.to_excel(writer, sheet_name=f'status_sheet{tmp}', index=False)
                    writer.save()
                    print(k)
                    break
            except exceptions.NoSuchElementException as e:
                browser.implicitly_wait(2)

        #Chat_results.csv
        chats = browser.find_elements(By.XPATH, "//span[contains(@class,'msg-text')]")
        t.sleep(0.5)
        datetime_stamp = datetime.datetime.now()
        timestamp = "Timestamp: %s-%s-%s %s:%s:%s" % (datetime_stamp.year, datetime_stamp.month, datetime_stamp.day, datetime_stamp.hour, datetime_stamp.minute, datetime_stamp.second )
        question_list = [conv_no,timestamp]
        reply_list = ['','','']
        for item in chats:
            chat_in_textformat = item.text
            if chat_in_textformat not in sent_replies_list:
                question_list.append(chat_in_textformat)
            else:
                reply_list.append(chat_in_textformat)
        question_list.append('')
        reply_list.append('')
        reply_list.append('')
        print(question_list)
        print(reply_list)
        print(time_logs_list)
        result_df['Questions'] = pd.Series(question_list)
        result_df['Replies'] = pd.Series(reply_list)
        result_df['Timelog'] = pd.Series(time_logs_list)
        result_df.to_csv('chat_results.csv', mode='a', index=False)
        conv_no = conv_no + 1
        browser.implicitly_wait(2)
        browser.refresh()
    # else:
    #     print('There seems to be an error in TimeStamp')


automation_script("response_with_status_temp.xlsx", "reschedule_sheet_small.xlsx", len_rows-3, len_rows-2)