import pandas as pd
import matplotlib.pyplot as plt
import sys
print("--------------------Wellcome to our Attendence Monitoring System----------------------")
while True: 
    name_file="name.xlsx"
    name_data=pd.read_excel(name_file)
    num_rows = name_data.shape[0]
    num_cols = name_data.shape[1]
    attendance_file1 = "OT.xlsx"
    attendance_data1 = pd.read_excel(attendance_file1)
    num_row_ot=attendance_data1.shape[0]
    num_col_ot=attendance_data1.shape[1]
    attendance_file2 = "python.xlsx"
    attendance_data2 = pd.read_excel(attendance_file2)
    num_row_py=attendance_data2.shape[0]
    num_col_py=attendance_data2.shape[1]
    attendance_file3 = "COA.xlsx"
    attendance_data3 = pd.read_excel(attendance_file3)
    num_row_coa=attendance_data3.shape[0]
    num_col_coa=attendance_data3.shape[1]
    attendance_file4 = "OS.xlsx"
    attendance_data4 = pd.read_excel(attendance_file4) 
    num_row_os=attendance_data4.shape[0] 
    num_col_os=attendance_data4.shape[1]  
    attendance_file5 = "java.xlsx"
    attendance_data5 = pd.read_excel(attendance_file5)
    num_row_java=attendance_data5.shape[0]
    num_col_java=attendance_data5.shape[1]#-------------- total Attendence of each subject ------------
    def sub_to(num_col_sub,attendance_data):
        total=0
        for i in range(3,num_col_sub,1):
            total=total+attendance_data.iloc[0][i]
        return total    #----------------------------------------------------
    ot_to=sub_to(num_col_ot,attendance_data1)   
    py_to=sub_to(num_col_py,attendance_data2)     
    coa_to=sub_to(num_col_coa,attendance_data3)   
    os_to=sub_to(num_col_os,attendance_data4)
    java_to=sub_to(num_col_java,attendance_data5) #---------- all subject total working days ------------   
    wr_days=(ot_to)+(py_to)+(coa_to)+(os_to)+(java_to)   
    in_uid=(input("Enter the full uid: ")) #-----------------UID check point----------------------  
    for i in range(0,num_rows,1):
        name = name_data.iloc[i][1]
        if(name==in_uid):
            count=1
            f=i
            break
        else:
            count=0       
    if(count!=1):
        print("It is not a valid UID!") 
        print("--------------------------------------------")
    else:      # ----------------Function for calculating attendence parsentage for each subject class-----------
        def sub_atten(attendance_data,f,num_col_sub): 
            subsum=0  
            for i in range(3,num_col_sub,1):
                subat=attendance_data.iloc[f][i]
                if(subat>0):
                    subsum=subsum+subat
            return subsum    
        ot_sum=sub_atten(attendance_data1,f,num_col_ot)
        py_sum=sub_atten(attendance_data2,f,num_col_py)
        coa_sum=sub_atten(attendance_data3,f,num_col_coa)
        os_sum=sub_atten(attendance_data4,f,num_col_os)
        java_sum=sub_atten(attendance_data5,f,num_col_java)
        total_atten=ot_sum+py_sum+coa_sum+os_sum+java_sum
        atten_p=(total_atten/wr_days)*100
        print("Name:",name_data.iloc[f][2])
        print(attendance_data1.iloc[f][2]," is present ",total_atten," days out of",wr_days,"days.")
        print("Overall Average Attendence: ",(total_atten/wr_days)*100)
        if(total_atten<60): # ------------------------------Eligiblity-----------------------------
            print("Candidate is not Eligible for Exam.")
        else:
            print("Candidate is Eligible for Exam.")
        ch=int(input("Want to know individual subject attendence parsentage?(Y=1/N=0): ")) #Individual subject attendence 
        if(ch==0):
            print("-----------ok-----------")
        else:
            print("Attendence in Optimization Technique is: ",ot_sum," and persentage is: ",((ot_sum/ot_to)*100))
            print("Attendence in Python is: ",py_sum," and persentage is: ",((py_sum/py_to)*100))
            print("Attendence in COA is: ",coa_sum," and persentage is: ",((coa_sum/coa_to)*100))
            print("Attendence in Operating System is: ",os_sum," and persentage is: ",((coa_sum/os_to)*100))
            print("Attendence in java is: ",java_sum," and persentage is: ",(java_sum/java_to)*100)
            print("--------------------------------------------")       
        choice=int(input("Want to see any graph?(Y=1/N=0): ")) # ---------------- Graph Moudule --------------------------- 
        if(choice==0):
            print("-----------ok-----------")
        else:
            print("----------Welcome to Graph Module----------")
            while(True):
                ch1=int(input("1.Attendence Days vs Subject 2. Attendence Persentsge Vs Subject 3. Exit from Graph Module : "))
                match ch1:
                    case 1: #------------------For choice 1 -------------------
                        print("----------Attendence Days vs Subject----------")
                        x1=["OT","Python","COA","OS","Java"]
                        y1=[ot_sum,py_sum,coa_sum,os_sum,java_sum]
                        plt.bar(x1,y1)
                        plt.xlabel("Subject Name")
                        plt.ylabel("Days")
                        plt.title("Attendence Days vs Subject")
                        plt.show()
                    case 2: #------------------For choice 2 -------------------
                        print("----------Attendence Persentsge Vs Subject----------")
                        x1=["OT","Python","COA","OS","Java"]
                        y1=[((ot_sum/java_to)*100),((py_sum/java_to)*100),((coa_sum/java_to)*100),((os_sum/os_to)*100),((java_sum/java_to)*100)]
                        plt.bar(x1,y1)
                        plt.xlabel("Subject-->")
                        plt.ylabel("Persentage-->")
                        plt.title("Attendence Persentsge Vs Subject")
                        plt.show()              
                    case 3:
                        break;  
        n=int(input("Want to Exit from program?(Y=1/N=0)?: "))
        if(n==1):
            print("--------------- Thank You --- Visit Again !--------------")
            sys.exit()
        else:
            print("-------------------------------------------------------")    