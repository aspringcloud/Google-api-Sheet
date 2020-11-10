import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = [
'https://spreadsheets.google.com/feeds',
'https://www.googleapis.com/auth/drive',
]

json_file_name = "impressive-bay-293705-d7e9899b0a5e.json"

credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
gc = gspread.authorize(credentials)


spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1srdgwIUAJRAxtiJWfSKS9p9GK3oag8JJMImI_JiC-s0/edit#gid=1817834757'

# 스프레스시트 문서 가져오기 
doc = gc.open_by_url(spreadsheet_url)

# 시트 선택하기
data_sheet = doc.worksheet('설문지 응답 시트1')
one_sheet = doc.worksheet('군산1호')
two_sheet = doc.worksheet('군산2호')
three_sheet = doc.worksheet('군산3호')
fore_sheet = doc.worksheet('군산4호')

# row 초기화
one_sheet.add_rows(999)
two_sheet.add_rows(999)
three_sheet.add_rows(999)
fore_sheet.add_rows(999)
one_sheet.delete_rows(2,999)
two_sheet.delete_rows(2,999)
three_sheet.delete_rows(2,999)
fore_sheet.delete_rows(2,999)

# 변수 선언
start_sheet_number = 4
start_car_one_sheet_number = 2
start_car_two_sheet_number = 2
start_car_three_sheet_number = 2
start_car_fore_sheet_number = 2
car_one_day_len_value = 0
car_two_day_len_value = 0
car_three_day_len_value = 0
car_fore_day_len_value = 0
temp_car_one_day_len_value = 0
temp_car_two_day_len_value = 0
temp_car_three_day_len_value = 0
temp_car_fore_day_len_value = 0
car_one_calculation_len = False
car_two_calculation_len = False
car_three_calculation_len = False
car_fore_calculation_len = False

cell_data = data_sheet.acell(f'A{start_sheet_number}').value
car_number = ""

def error_content(row_number):
    result = ""
    if row_number == 7:
        result = "A"
    elif row_number == 8:
        result = "B"
    elif row_number == 9:
        result = "C"
    elif row_number == 10:
        result = "D"
    elif row_number == 11:
        result = "E"
    elif row_number == 12:
        result = "F"
    elif row_number == 13:
        result = "G"
    elif row_number == 14:
        result = "H"
    elif row_number == 15:
        result = "I"
    elif row_number == 16:
        result = "J"
    elif row_number == 0:
        result = "X"

    return result

def create_row(sheet,data,row_number,day_len):
    '''
        data[0] = 타임스템프 
        data[4] = 운행시작시간 
        data[5] = 운행종료시간 
        day_len = 일일 운행거리(계산)
        data[2] = 안전요원 
        data[21]+data[22] = 금일날씨 + 온도 
        data[20] = 수동조작 횟수 
        error_content(row_number) = 오류코드
        data[23] = DTG
        data[24] = DVR
    '''
    sheet.append_row([data[0],data[4],data[5],day_len,data[6],data[2],"?",data[21]+data[22],"?",data[20],error_content(row_number),data[23],data[24]], start_car_two_sheet_number)
# 행 생성 반복문
while(cell_data != ""):
    # 변수선언
    cell_data = data_sheet.acell(f'A{start_sheet_number}').value
    row_data = data_sheet.row_values(start_sheet_number)
    # print(row_data)
    # 행 만드는 로직
    try:
        if row_data[1] == "임1146":

            if car_one_calculation_len:
                car_one_day_len_value = int(row_data[6]) - int(temp_car_one_day_len_value)
                temp_car_one_day_len_value = int(row_data[6])
            else:
                temp_car_one_day_len_value = int(row_data[6])
                car_one_calculation_len = True

            row_change_number = 7
            count_error = 0
            
            while(row_change_number < 17):
                if row_data[row_change_number] != "":
                    create_row(one_sheet,row_data,row_change_number,car_one_day_len_value)
                    count_error += 1
                
                row_change_number += 1

            if count_error == 0:
                create_row(one_sheet,row_data,count_error,car_one_day_len_value)
            
            start_car_one_sheet_number += 1
        elif row_data[1] == "임1147":

            if car_two_calculation_len:
                car_two_day_len_value = int(row_data[6]) - int(temp_car_two_day_len_value)
                temp_car_two_day_len_value = int(row_data[6])
            else:
                temp_car_two_day_len_value = int(row_data[6])
                car_two_calculation_len = True

            row_change_number = 7
            count_error = 0
            
            while(row_change_number < 17):
                if row_data[row_change_number] != "":
                    create_row(two_sheet,row_data,row_change_number,car_two_day_len_value)
                    count_error += 1
                row_change_number += 1

            if count_error == 0:
                create_row(two_sheet,row_data,count_error,car_two_day_len_value)

            start_car_two_sheet_number += 1
            car_two_day_len_value = int(row_data[6]) - int(car_two_day_len_value) 
        elif row_data[1] == "임6894":

            if car_three_calculation_len:
                car_three_day_len_value = int(row_data[6]) - int(temp_car_three_day_len_value)
                temp_car_three_day_len_value = int(row_data[6])
            else:
                temp_car_three_day_len_value = int(row_data[6])
                car_three_calculation_len = True

            row_change_number = 7
            count_error = 0
            
            while(row_change_number < 17):
                if row_data[row_change_number] != "":
                    create_row(three_sheet,row_data,row_change_number,car_three_day_len_value)
                    count_error += 1
                row_change_number += 1

            if count_error == 0:
                create_row(three_sheet,row_data,count_error,car_three_day_len_value)

            start_car_three_sheet_number += 1
            car_three_day_len_value = int(row_data[6]) - int(car_three_day_len_value) 
        elif row_data[1] == "임6895":

            if car_fore_calculation_len:
                car_fore_day_len_value = int(row_data[6]) - int(temp_car_fore_day_len_value)
                temp_car_fore_day_len_value = int(row_data[6])
            else:
                temp_car_fore_day_len_value = int(row_data[6])
                car_fore_calculation_len = True

            row_change_number = 7
            count_error = 0
            
            while(row_change_number < 17):
                if row_data[row_change_number] != "":
                    create_row(fore_sheet,row_data,row_change_number,car_fore_day_len_value)
                    count_error += 1
                row_change_number += 1

            if count_error == 0:
                create_row(fore_sheet,row_data,count_error,car_fore_day_len_value)

            start_car_fore_sheet_number += 1
            car_fore_day_len_value = int(row_data[6]) - int(car_fore_day_len_value) 
    except Exception as e:
        print(e)
    

    start_sheet_number += 1 