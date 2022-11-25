# from asyncio.windows_events import NULL
# from enum import Flag
from pickle import TRUE
import openpyxl
import os
action_arr=[]
expected_arr=[]
random_arr=[]
testcase=''
action_text=''
xml1=''
xml_final_text=''
files_name=[]
files = [f for f in os.listdir('.') if os.path.isfile(f)]
for f in files:
    fname=f.endswith('.xlsx')
    # print(fname)
    if(fname):
        files_name.append(f)
 
# print(files_name)
excel_file_name=files_name[0]
wb_obj = openpyxl.load_workbook(excel_file_name) 

sheet = wb_obj.active
# print(sheet['G2'].value)
maxnum=0
no_of_testcase=0
for i in range(100):
    
    txt=''
    txt=str(sheet[f'B{i+1}'].value)
    # print(txt)
    if(len(txt)>4):
        maxnum=maxnum+1
    else:
        break
# print("maxnum is:::::::::",maxnum)



flag=False
arr=[]
# print("maxnum is",maxnum)
for row in sheet.iter_rows(max_row=maxnum):
        no_of_testcase+=1
        for cell in row:
            if(flag):
                arr.append(cell.value)        
        action_text=''
        if(flag): 
            Test_Objective=arr[0]
            Module=arr[1]
            User_Story=arr[2]
            User_Story=User_Story.strip()
            
            # print(Test_Objective)
            # print(Module)
            # print(User_Story)
            AC_Mapping=arr[3]
            Importance=arr[4]
            print('importance is:',Importance)
            Importance=str(Importance.strip())
            Importance=Importance.lower()
            if(Importance=="medium"):
                Importance=2
            elif(Importance=="low"):
                Importance=1
            elif(Importance=="high"):
                Importance=3
            else:
                print("something wrong with Importance")
            # print(Importance)
            Pre_Requisite=arr[5]
            precondition_text=''
            prequisite_flag=False
            for i in range(1,100):
                if(f"{i}." in Pre_Requisite):
                    # print("Found in 2nd elif")
                    # print('i am in jjjjjjjj')
                    pre_index=Pre_Requisite.find(f"{i}.")
                    pre_index2=Pre_Requisite.find(f"{i+1}.")
                    # print("index values is ",index,index2)
                    if(pre_index2==-1):
                        pre_text1=Pre_Requisite[pre_index:]
                        prequisite_flag=True

                        # test_steps=test_steps.replace(f'Step {k}',"")
                    else:
                        pre_text1=Pre_Requisite[pre_index:pre_index2]
                        # print("text is:::::",text1)
                    pre_text1=pre_text1.strip()
                    if(prequisite_flag):
                        precondition_text=precondition_text+pre_text1
                    else:
                        precondition_text=precondition_text+pre_text1+f"<br/>"
                else:
                    precondition_text=precondition_text+Pre_Requisite
                    prequisite_flag=True


                if(prequisite_flag):
                    break
                


            Actions=arr[6]

            Expected_Result=arr[7]
            # print('actions are ',Actions)
            # print("Action is: ",Actions)
            # print('hhhhhhhhhhh')
            try:
                index_initial=int(Actions.index('Initial Steps'))
                # print("intial indec is ",index_initial)
            except:
                index_initial='not found'
            try:
                index2_initial=int(Actions.index('Test Steps'))
            except:
                index2_initial='not found'
            # print("index number is",index2_initial)
            # print('indexes are::',Actions[index2_initial:])
            
            if(index_initial=='not found' or index2_initial == 'not found'):
                print('indexes not found in action steps')
            else:
                initial_step=Actions[index_initial:index2_initial-1]
                # Actions=Actions.replace(initial_step,"")
                test_steps=Actions[index2_initial:]
                # print(test_steps)
                randomtext=''
                randomtext=initial_step
                randomtext=randomtext.replace('Initial Steps:',"")
                randomtext=randomtext.replace('-',"")
                action_text=action_text+f'''
            <step>

                    <step_number><![CDATA[{1}]]></step_number>
                    <actions><![CDATA[<b>Initial Step:-</b><br/>{randomtext}]]>
                    </actions>

                    <expectedresults><![CDATA[Success]]>
                    </expectedresults>
                    <execution_type><![CDATA[1]]></execution_type>
            
                </step>'''
                # print(randomtext)

                for l in range(1,100):
                    if(f"{l}." in randomtext):
                        # print("Found in 2nd elif")
                        # print('i am in jjjjjjjj')
                        random_index=randomtext.find(f"{l}.")
    
                        random_index2=randomtext.find(f"{l+1}.")
                        if(random_index2==-1):
                            text1=randomtext[random_index:]
                            
                            # test_steps=test_steps.replace(f'Step {k}',"")
                        else:
                            text1=randomtext[random_index:random_index2]
                            # print("text is:::::",text1)
                        text1=text1.strip()
                        random_arr.append(text1)
                # print(len(random_arr))

                flag=False
                flag2=False
                for k in range(len(random_arr),100):
                    # test_ind=test_steps.find(f"{k}.")
                    # textStepText=test_steps[test_ind-1:test_ind]
                    # print("tesssssss:",textStepText)

                    if(f"{k}." in test_steps):
                        # print("Found in 2nd elif")
                        # print('i am in jjjjjjjj')
                        index = test_steps.find(f"{k}.")
                        index2=test_steps.find(f"{k+1}.")
                        # print("index values is ",index,index2)
                        if(index2==-1):
                            text1=test_steps[index:]
                            flag=True
                            # test_steps=test_steps.replace(f'Step {k}',"")
                        else:
                            text1=test_steps[index:index2]
                            # print("text is:::::",text1)
                        text1=text1.strip()
                        text1=text1.replace(f"{k}.","")
                        text1=text1.strip()
                        # print(text1)
                        action_arr.append(text1)
                        # test_steps=test_steps.replace(f'Step {k}',"")
                        # print("text is:",text1)
                        # print("text is:",text1)
                    if(flag):
                        break
                # print(action_arr)
                no_slice=False
                for x in range(1,100):
                    if(f"{x}." in Expected_Result):
                        # print("Found in 2nd elif")
                        index = Expected_Result.find(f"{x}.")
                        index2=Expected_Result.find(f"{x+1}.")
                        # print("index values is ",index,index2)
                        if(index2==-1):
                            text1=Expected_Result[index:]
                            flag2=True
                        else:
                            text1=Expected_Result[index:index2]
                        text1=text1.strip()
                        expected_arr.append(text1)
                        # Expected_Result=Expected_Result.replace(f'Step {k}',"")
                    
                    if(flag2):
                        break

                if(flag2!=True):
                    expected_arr.append(Expected_Result)
                    no_slice=True
                    
                # print(action_arr)
                # print(expected_arr)


                xml2=''
                xml2=xml2+f''' 
    <testcase internalid="" name="{User_Story}">
        <node_order><![CDATA[]]></node_order>
        <externalid><![CDATA[]]></externalid>
        <version><![CDATA[]]></version>
        <summary><![CDATA[{Test_Objective}]]></summary>
        <preconditions><![CDATA[{precondition_text}]]></preconditions>
        <execution_type><![CDATA[1]]></execution_type>
        <importance><![CDATA[{Importance}]]></importance>
        <steps>'''
                # print(len(action_arr))
                # print(len(expected_arr))
                j=0
                t=2
                num=0
                num=len(action_arr)-len(expected_arr)
                for i in range(len(action_arr)):
                    try:
                        a=action_arr[i]
                    except:
                        break
                    try:
                        b=expected_arr[j]   
                    except:
                        break
                    # print("a is",a,"\nb is",b)                                                                                                                                                                                   
                    # a=a[0:6]
                    # a=a.strip()
                    # b=a[0:6]
                    # index_a= a.find(f"{i+1}")
                    # index_b=b.find(f"Step")
                    # print("index values is ",index,index2)
                    a=a.strip()
                    b=b.strip()
                    if(no_slice):
                        b=b.strip()

                    else:
                        text_b=b[0:2]
                        b=b.replace(text_b,"")
                        b=b.strip()
                       
                        
                    
                    if(i>=num):
                        action_text=action_text+f'''
            <step>

                    <step_number><![CDATA[{t}]]></step_number>
                    <actions><![CDATA[{a}]]>
                    </actions>

                        <expectedresults><![CDATA[{b}]]>
                        </expectedresults>
                        <execution_type><![CDATA[1]]></execution_type>
                
                    </step>'''
                        t=t+1
                        j=j+1

                    else:
                        action_text=action_text+f'''
            <step>

                    <step_number><![CDATA[{t}]]></step_number>
                    <actions><![CDATA[{a}]]>
                    </actions>

                    <expectedresults><![CDATA[Success]]>
                    </expectedresults>
                    <execution_type><![CDATA[1]]></execution_type>
            
                </step>'''
                        t=t+1
                    


                xml_final_text=f'''    </steps>

        

            </testcase>

            '''
                testcase=testcase+xml2+action_text+xml_final_text
                # print(testcase)                
                # print("text is:",text1)

            arr=[]
            action_arr=[]
            expected_arr=[]
            random_arr=[]
        flag=True 

start_text=f'''<?xml version="1.0" encoding="UTF-8"?>

<testcases>

'''

end_text=f'''



</testcases>'''


xml_file_text=start_text+testcase+end_text
# print(xml_file_text)
with open('readme.xml', 'w') as f:
    f.write(xml_file_text)

print("XML CREATED SUCCESSFULLY")
print("Number Of Testcases is:",no_of_testcase-1)
input()

# For expected Result you can add the data by removing the number and extra spaces and make changes accordance in last this 
# will be more benefecial