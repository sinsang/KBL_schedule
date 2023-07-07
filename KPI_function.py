import openpyxl as xl

# 팀 이름
team = [
    "",
    "SK",
    "삼성",
    "고양",
    "대구",
    "안양",
    "원주",
    "전주",
    "울산",
    "수원",
    "창원",
    ]

# 팀별 경기 저장
# 휴식0홈1원정2, 휴일인지여부, 상대팀  
team_game = [
    [],
    [[0, False, -1] for i in range(177)],
    [[0, False, -1] for i in range(177)],
    [[0, False, -1] for i in range(177)],
    [[0, False, -1] for i in range(177)],
    [[0, False, -1] for i in range(177)],
    [[0, False, -1] for i in range(177)],
    [[0, False, -1] for i in range(177)],
    [[0, False, -1] for i in range(177)],
    [[0, False, -1] for i in range(177)],
    [[0, False, -1] for i in range(177)],
    ]

# 휴일 목록 
holidays = [1,2,8,9,15,16,22,23,43,44,50,51,57,58,64,65,71,72,78,79,92,93,99,100,106,107,113,114,134,135,148,149,155,156,162,163,169,170,66,73,112,115,133]

result = xl.load_workbook("결과.xlsx")
result_sheet = result["Sheet1"]

# 경기 저장
# 홈, 어웨이, day
games = []

print("경기 배치 결과 읽어들이는 중...")
for r in result_sheet.rows:

    tmp = []

    for c in r:
        try:
            tmp.append(int(c.value))
        except:
            break
        
    if len(tmp) > 0:
        games.append(tmp)

result.close()
# 경기 저장 끝

# 경기 정렬
games = sorted(games, key=lambda x:x[2])

# 지표 탐색
new_wb = xl.load_workbook("KPI2.xlsx", data_only=True)
new_ws = new_wb["Sheet1"]

for g in games:

    home = g[0]
    away = g[1]
    day = g[2] - 1

    team_game[home][day][0] = 1
    team_game[away][day][0] = 2
    team_game[home][day][2] = away
    team_game[away][day][2] = home
    if (day+1) in holidays:
        team_game[home][day][1] = True
        team_game[away][day][1] = True
print("경기 배치 결과 저장 완료")

# 최대 홈 연속
row = 3
for tg in team_game[1:]:
    
    day = 1

    result = 0
    resultStartDay = 0
    resultEndDay = 0
    
    startDay = 0
    count = 0
    
    for g in tg:
        if g[0] == 1:
            if count == 0:
                startDay = day
            count += 1
        elif g[0] == 2:
            count = 0
        if result < count:
            result = count
            resultStartDay = startDay
            resultEndDay = day
        day += 1
        
    #print(result, resultStartDay, resultEndDay)
    new_ws['J' + str(row)] = result
    row += 1
print("최대 홈 연속 계산 완료")
    
# 최대 원정 연속
row = 3
for tg in team_game[1:]:
    
    day = 1

    result = 0
    resultStartDay = 0
    resultEndDay = 0
    
    startDay = 0
    count = 0
    
    for g in tg:
        if g[0] == 2:
            if count == 0:
                startDay = day
            count += 1
        elif g[0] == 1:
            count = 0
        if result < count:
            result = count
            resultStartDay = startDay
            resultEndDay = day
        day += 1

    #print(result, resultStartDay, resultEndDay)
    new_ws['K' + str(row)] = result
    row += 1
print("최대 원정 연속 계산 완료")

# 휴일 경기 수
row = 3
for tg in team_game[1:]:
    result = 0
    
    for g in tg:
        if g[1]:
            result += 1
            
    new_ws['B' + str(row)] = result
    row += 1
print("휴일 경기수 계산 완료")

# 연전 횟수
row = 3
for tg in team_game[1:]:
    result = 0

    for g in range(len(tg)-1):
        if tg[g][0] != 0 and tg[g+1][1] != 0:
            result += 1
            
    new_ws['C' + str(row)] = result
    row += 1
print("연전 횟수 계산 완료")


# 퐁당
row = 3
pong = []
for tg in team_game[1:]:
    result4 = 0
    result5 = 0
    result6 = 0

    g = 0
    count = 0
    while g < len(tg)-2:

        if tg[g][0] != 0 and tg[g+1][0] == 0 and tg[g+2][0] != 0:
            count += 1
            g += 2
        else:
            if count == 3:
                result4 += 1
            elif count == 4:
                result5 += 1
            elif count >= 5:
                result6 += 1
            count = 0
        
            g += 1

    new_ws['G' + str(row)] = result4
    new_ws['H' + str(row)] = result5
    new_ws['I' + str(row)] = result6
    pong.append((result4, result5, result6))
    row += 1
print("퐁당 배정 계산 완료")

# 7일 4경기
row = 3
now = 0
for tg in team_game[1:]:

    result = 0

    #print(now)
    for g in range(len(tg)-6):
        count = 0
        flag = True
        if tg[g][0] != 0 and tg[g+6][0] != 0:
            for i in range(7):
                if count >= 4:
                    flag = False
                    break
                if tg[g+i][0] != 0:
                    count += 1

            if count >= 4 and flag:
                result += 1
                #print(*tg[g:g+7])

    result -= pong[now][0] + 2*pong[now][1] + 3*pong[now][2]
    now += 1
    if result < 0:
        result = 0
    new_ws['D' + str(row)] = result
    row += 1
    
print("7일 4경기 계산 완료")

# 6일 4경기
row = 3
_6_4 = []
for tg in team_game[1:]:
    result = 0
    for g in range(5, len(tg)-5, 7):
        count = 0
        for i in range(6):
            if tg[g+i][0] != 0:
                count += 1

        if count >= 4:
            result += 1
    _6_4.append(result)
    new_ws['F' + str(row)] = result
    row += 1
print("6일 4경기 계산 완료")

# 4일 3경기
row = 3
now = 0
for tg in team_game[1:]:
    result = 0
    
    # 목토일 
    for g in range(5, len(tg)-3, 7):
        count = 0
        for i in range(4):
            if tg[g+i][0] != 0:
                count += 1

        if count >= 3:
            result += 1

    # 토일화
    for g in range(0,len(tg)-3, 7):
        count = 0
        for i in range(4):
            if tg[g+i][0] != 0:
                count += 1
        if count >= 3:
            result += 1

    #print()
    result -= 2*_6_4[now]
    if result < 0:
        result = 0

    now += 1
    new_ws['E' + str(row)] = result
    row += 1
print("4일 3경기 계산 완료")

# 라운드별 홈 경기
'''
r = [(1,23),
    (24,58),
    (59,89),
    (90,113),
    (114,149),
    (150,172),
     ]

row = 3
for tg in team_game[1:]:
    col = 'L'
    for i, j in r:
        count = 0
        for g in tg[i-1: j]:
            if g[0] == 1:
                count += 1
        new_ws[col + str(row)] = count
        col = chr(ord(col) + 1)
    row += 1
'''

r = []

row = 3
now = 0
for tg in team_game[1:]:
    col = 'L'
    roundCount = 0
    roundCheck = [False for i in range(10)]
    roundCheck[now] = True
    count = 0

    for g in tg:
        if g[0] == 0:
            continue
        roundCheck[g[2]-1] = True
        #print(*roundCheck, count)
        if g[0] != 0:
            count += 1
        if not (False in roundCheck):
            new_ws[col + str(row)] = count
            col = chr(ord(col) + 1)
            roundCount += 1
            count = 0
            roundCheck = [False for i in range(10)]
            roundCheck[now] = True
    print(now, roundCount)
    row += 1
    now += 1
            

print("라운드별 홈 경기 계산 완료")

# 휴일 홈 경기 수
row = 3
for tg in team_game[1:]:
    result = 0
    
    for g in tg:
        if g[1] and g[0] == 1:
            result += 1
            
    new_ws['R' + str(row)] = result
    row += 1
print("휴일 홈 경기수 계산 완료")



new_wb.save("KPI_result.xlsx")
new_wb.close()



