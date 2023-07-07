import openpyxl as xl

# 결과정보를 출력할지
infoPrint = False

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
# 휴식0홈1원정2슈퍼리그3, 휴일인지여부, 상대팀  
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

# 슈퍼리그 일정 추가
team_game[5][5][0] = 3
team_game[5][47][0] = 3
team_game[5][54][0] = 3
team_game[5][82][0] = 3
team_game[5][96][0] = 3
team_game[5][110][0] = 3
team_game[5][140][0] = 3
team_game[5][142][0] = 3

team_game[1][12][0] = 3
team_game[1][19][0] = 3
team_game[1][68][0] = 3
team_game[1][103][0] = 3
team_game[1][110][0] = 3
team_game[1][140][0] = 3
team_game[1][142][0] = 3

# KPI 계산 함수 : 적합 부적합 판정
def calcKPI (sch):

    # 경기 저장
    # 홈, 어웨이, day
    games = []
    d = 1

    for r in sch:
        for c in r:
            games.append([c[0], c[1], d])
        d += 1

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

    # KPI 결과값
    RESULT = {
        'holidayGame' : [0 for i in range(11)],     # 휴일 경기수
        'steak' : [0 for i in range(11)],           # 연전횟수
        '7_4' : [0 for i in range(11)],             # 7일 4경기
        '4_3' : [0 for i in range(11)],             # 4일 3경기
        '6_4' : [0 for i in range(11)],             # 6일 4경기
        'pong4' : [0 for i in range(11)],           # 퐁당 4연속
        'pong5' : [0 for i in range(11)],           # 퐁당 5연속
        'pong6' : [0 for i in range(11)],           # 퐁당 6연속
        'home' : [0 for i in range(11)],            # 최대 홈 연속
        'away' : [0 for i in range(11)],            # 최대 어웨이 연속
        'roundHome' : [[0 for i in range(6)] for j in range(11)],    # 라운드별 홈 경기
        'holidayHome' : [0 for i in range(11)],
    }

    # KPI 계산
    # 최대 홈 연속
    teamCount = 0
    for tg in team_game[1:]:

        teamCount += 1
        day = 1
        result = 0

        count = 0
        
        for g in tg:
            if g[0] == 1:
                count += 1

            elif g[0] == 2:
                count = 0
            if result < count:
                result = count

            day += 1

        RESULT['home'][teamCount] = result
    # 최대 홈 연속 계산 완료

    # 최대 원정 연속
    teamCount = 0
    for tg in team_game[1:]:

        teamCount += 1
        day = 1
        result = 0
        count = 0
        
        for g in tg:
            if g[0] == 2:
                count += 1
            elif g[0] == 1:
                count = 0
            if result < count:
                result = count

            day += 1

        RESULT['away'][teamCount] = result
    # 최대 원정 연속 계산 완료

    # 휴일 경기 수
    teamCount = 0
    for tg in team_game[1:]:
        teamCount += 1
        result = 0
        
        for g in tg:
            if g[1]:
                result += 1
        
        RESULT['holidayGame'][teamCount] = result
    # 휴일 경기수 계산 완료

    # 연전 횟수
    teamCount = 0
    for tg in team_game[1:]:
        teamCount += 1
        result = 0

        for g in range(len(tg)-1):
            if tg[g][0] != 0 and tg[g+1][1] != 0:
                result += 1
                
        RESULT['steak'][teamCount] = result
    # 연전 횟수 계산 완료

    # 퐁당
    pong = []
    teamCount = 0
    for tg in team_game[1:]:

        teamCount += 1
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

        RESULT["pong4"][teamCount] = result4
        RESULT["pong5"][teamCount] = result5
        RESULT["pong6"][teamCount] = result6

        pong.append((result4, result5, result6))
    # 퐁당 배정 계산 완료

    # 7일 4경기
    now = 0
    teamCount = 0
    for tg in team_game[1:]:

        result = 0
        teamCount += 1
        
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

        result -= pong[now][0] + 2*pong[now][1] + 3*pong[now][2]
        now += 1
        if result < 0:
            result = 0

        RESULT['7_4'][teamCount] = result
    # 7일 4경기 계산 완료

    # 6일 4경기
    teamCount = 0
    _6_4 = []
    for tg in team_game[1:]:

        teamCount += 1
        result = 0

        for g in range(5, len(tg)-5, 7):
            count = 0
            for i in range(6):
                if tg[g+i][0] != 0:
                    count += 1

            if count >= 4:
                result += 1
                
        _6_4.append(result)
        RESULT['6_4'][teamCount] = result
    # 6일 4경기 계산 완료

    # 4일 3경기
    now = 0
    teamCount = 0
    for tg in team_game[1:]:

        teamCount += 1
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

        result -= 2*_6_4[now]
        if result < 0:
            result = 0
        
        RESULT['4_3'][teamCount] = result

        now += 1
    # 4일 3경기 계산 완료

    # 라운드별 홈 경기 수 계산
    now = 0
    teamCount = 0
    for tg in team_game[1:]:
        
        teamCount += 1
        roundCount = 0
        roundCheck = [False for i in range(10)]
        roundCheck[now] = True
        count = 0

        index = 1
        for g in tg:
        
            index += 1

            if g[0] == 0 or g[0] == 3:
                continue

            roundCheck[g[2]-1] = True

            if g[0] == 1:
                count += 1

            if not (False in roundCheck):
                RESULT['roundHome'][teamCount][roundCount] = count
                roundCount += 1
                count = 0
                roundCheck = [False for i in range(10)]
                roundCheck[now] = True

        now += 1
    # 라운드별 홈 경기 계산 완료

    # 휴일 홈 경기 수
    teamCount = 0
    for tg in team_game[1:]:

        teamCount += 1
        result = 0
        
        for g in tg:
            if g[1] and g[0] == 1:
                result += 1
        RESULT['holidayHome'][teamCount] = result
    #휴일 홈 경기수 계산 완료

    # KPI 결과값 검증
    for i in RESULT:
        if i == 'roundHome':
            continue
        if max(RESULT[i][1:]) - min(RESULT[i][1:]) > 2:
            return False
    
    return True
