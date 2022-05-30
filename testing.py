import time as t
from selenium import webdriver
from selenium.common import exceptions
import pandas as pd
from selenium.webdriver.common.by import By
import datetime


browser = webdriver.Firefox()
browser.maximize_window()
browser.get('https://chat.rootle.ai/')
browser.implicitly_wait(3)

all_sheets_dict = pd.read_excel('response_sheet.xlsx', sheet_name=None)
reschedule_sheet = pd.read_excel("reschedule_sheet.xlsx")
no_of_rows = all_sheets_dict['Sheet1'].index
sent_replies = []
reschedule_index_increment = 0
writer = pd.ExcelWriter("response_with_status.xlsx", engine = 'openpyxl', mode = 'a', if_sheet_exists = 'overlay') # pylint: disable=abstract-class-instantiated
reschedule_writer = pd.ExcelWriter("reschedule_sheet.xlsx", engine = 'openpyxl', mode = 'a', if_sheet_exists = 'overlay') # pylint: disable=abstract-class-instantiated

#Check greeting message on the basis of time-stamp.
time_stamp_flag = False
greeting_msg = browser.find_element_by_class_name('msg-text').text
am_or_pm = browser.find_element_by_class_name('time-stamp').text.split(' ')
if am_or_pm[1] == 'pm':
    time = int(am_or_pm[0].split(':')[0]) + 12
    if time>=11 and time<=17 or time == 24:
        if greeting_msg in ['Hello John, Very Good afternoon.', 'Hello John, Good afternoon.']:
            time_stamp_flag = True
    if time>=18 and time<=23:
        if greeting_msg in ['Hello John, Very Good evening.', 'Hello John, Good evening.']:
            time_stamp_flag=True
elif am_or_pm[1] == 'am':
    if greeting_msg in ['Hello John, Very Good morning.', 'Hello John, Good morning.']:
        time_stamp_flag = True

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
                    'I am calling from Dell Technologies. Is it right time to talk?',
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
if time_stamp_flag:
    for i in range(len(no_of_rows)):
        print(i,'++++++++++')
        count = 0
        result_df = pd.DataFrame(columns=['Questions','Replies'])
        flag = False
        tmp = 1
        chat_history_list = []

        for k,v in all_sheets_dict.items():
            k = pd.DataFrame(v)
            try:
                reply_field = browser.find_element_by_xpath('/html/body/div/div/div/div[2]/div/input')
                reply_field.send_keys(k.iloc[i,1])
                reply_button = browser.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/button')
                reply_button.click()
                t.sleep(1)
                sent_replies.append(k.iloc[i,1])

                first_ques = browser.find_element_by_xpath("/html/body/div/div/div/div[1]/div[2]/div[1]/div/div/span").text
                all_element_list = browser.find_elements_by_class_name("msg-text")
                
                for chat in all_element_list:
                    current_chat = chat.text

                if current_chat in gibberish_question_list:
                    k.iloc[i,2] = 'Failed'
                    flag = True
                elif current_chat == first_ques or current_chat in chat_history_list:
                    k.iloc[i,2] = 'Repeated'
                    flag=True
                    break
                elif current_chat in pass_question_list:
                    k.iloc[i,2] = 'Passed'
                elif current_chat in reschedule_question_list:
                    k.iloc[i,2] = 'Rescheduled'
                    for j in range(len(reschedule_sheet.index)):
                        reply_field = browser.find_element_by_xpath('/html/body/div/div/div/div[2]/div/input')
                        reply_field.send_keys(reschedule_sheet.iloc[j+reschedule_index_increment,1])
                        reply_button = browser.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/button')
                        reply_button.click()

                        all_element_list_for_rescheduling = browser.find_elements_by_class_name("msg-text")
                        
                        for chat in all_element_list_for_rescheduling:
                            current_chat_for_rescheduling = chat.text
                        if current_chat_for_rescheduling in regards_list:
                            reschedule_sheet.iloc[j+reschedule_index_increment,2] = 'Passed'
                        reschedule_sheet.to_excel(reschedule_writer, index=False)
                        reschedule_writer.save()
                        t.sleep(1)

                        sent_replies.append(reschedule_sheet.iloc[j+reschedule_index_increment,1])
                        reschedule_index_increment = reschedule_index_increment + 1
                        flag = True
                        break
                print(k)
                k.to_excel(writer, sheet_name=f'status_sheet{tmp}', index=False)
                tmp = tmp + 1
                writer.save()

                if flag:
                    flag = False
                    break
                chat_history_list.append(current_chat)
            except exceptions.NoSuchElementException as e:
                browser.implicitly_wait(2)

        chats = browser.find_elements(By.XPATH, "//span[contains(@class,'msg-text')]")
        datetime_stamp = datetime.datetime.now()
        edited = "Timestamp: %s-%s-%s %s:%s:%s" % (datetime_stamp.year, datetime_stamp.month, datetime_stamp.day, datetime_stamp.hour, datetime_stamp.minute, datetime_stamp.second )
        question_list = [edited]
        reply_list = ['','']
        for item in chats:
            chat_in_textformat = item.text
            if chat_in_textformat not in sent_replies:
                question_list.append(chat_in_textformat)
            else:
                reply_list.append(chat_in_textformat)
        question_list.append('')
        reply_list.append('')
        reply_list.append('')
        print(reply_list)
        print(question_list)
        result_df['Questions'] = pd.Series(question_list)
        result_df['Replies'] = pd.Series(reply_list)
        result_df.to_csv('chat_results.csv', mode='a', index=False)

        # zip_obj = zip(question_list,reply_list)
        # zipped_list = list(zip_obj)
        # chat_df =  pd.DataFrame(zipped_list, columns=['Questions','Replies'])
        # print(chat_df)
        # writer = pd.ExcelWriter("results.xlsx" , engine='openpyxl', mode='a' ) # pylint: disable=abstract-class-instantiated
        # chat_df.to_excel(writer, sheet_name='Results')
        # writer.save()

        browser.implicitly_wait(2)
        browser.refresh()
else:
    print('There seems to be an error in TimeStamp')
