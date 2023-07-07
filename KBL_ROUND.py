import openpyxl as xl
from collections import deque
import KPI_func
from itertools import combinations, permutations

wb = xl.load_workbook("결과.xlsx")

ws = wb['Sheet1']

matches = []

for r in ws:
    tmp = []
    for c in r[:2]:
        try:
            tmp.append(int(c.value))
        except:
            break
    if len(tmp) == 2:
        matches.append(tuple(tmp))

matches.sort(key=lambda x:x[0])
matches = list(set(matches))

matches_var1 = permutations(matches, 1)
print(len(list(matches_var1)))

matches_var2 = permutations(matches, 2)
print(len(list(matches_var2)))

matches_var3 = permutations(matches, 3)
print(len(list(matches_var3)))

matches_var = [[], matches_var1, matches_var2, matches_var3]

print(len(matches))

'''
 각 노드는 모든 경기의 배치 리스트
 다음 노드로 진출 시 KPI 계산
 KPI 계산 값이 기준에 부합하지 않거나, 기본 제약조건에 부합하지 않을 시 노드 진출 안함
'''

# 휴일 목록 
holidays = [1,2,8,9,15,16,22,23,43,44,50,51,57,58,64,65,71,72,78,79,92,93,99,100,106,107,113,114,134,135,148,149,155,156,162,163,169,170,66,73,112,115,133]

# 각 일정별 배치해야 하는 경기 수
mustGames = [3,3,
            1,2,0,2,2,3,3,
            1,2,0,2,2,3,3,
            1,2,0,2,2,3,3,
            1,2,2,2,0,0,0,
            0,0,0,0,0,0,0,
            0,0,2,2,2,3,3,
            1,2,0,2,2,3,3,
            1,2,0,2,2,3,3,
            1,1,1,2,2,3,3,
            3,1,0,1,2,3,3,
            3,1,1,1,1,3,3,
            2,2,0,0,0,0,0,
            0,0,2,2,2,3,3,
            1,2,0,2,2,3,3,
            1,2,0,2,2,3,3,
            1,2,0,2,2,3,3,
            2,2,2,2,0,0,0,
            0,0,0,0,0,0,0,
            0,0,2,2,3,3,3,
            1,1,2,2,0,2,0,
            1,1,1,2,2,3,3,
            1,1,1,2,2,3,3,
            1,1,1,2,2,3,3,
            1,1,1,2,2,3,3,
            0,5]
 
print(len(mustGames), sum(mustGames))
q = deque([])
q.append((0, []))   # 비어있는 0일차로 초기화 

# (현재 일자, [배치 결과])

while q:

    tmp = q.popleft()

    day = tmp[0]
    sch = tmp[1]

    #print("day", day)
    print(sch)

    #print(day)
    
    if day > 20:#len(mustGames)-1:
        print("성공")
        print(sch) 
        print()
        continue

    # 일자 별 경기만큼 배치
    mustGameCount = mustGames[day]  # 배치해야하는 게임의 수
    matchCount = dict()

    # 이미 배치된 경기의 매치카운트 계산
    for i in range(len(sch)):
        for j in range(len(sch[i])):
            try:
                matchCount[str(sch[i][j])] += 1
            except:
                matchCount[str(sch[i][j])] = 1

    matchTmp = []  # 배치 넣을 게임
    
    for i in range(len(matches)):

        nowMatch = matches[i]

        if mustGameCount == 0:
            tsch = sch[:]
            tsch.append(matchTmp)
            q.append([day + 1, tsch])
            break

        # 배치될 경기가 겹칠 경우 중단
        if nowMatch in matchTmp:
            continue

        # 배치될 경기의 팀이 당일날 있어도 중단
        flag = False
        for m in matchTmp:
            if m[0] in nowMatch or m[1] in nowMatch:
                flag = True
                break
        if flag:
            continue

        # 배치될 경기가 카운트를 넘을 경우 중단
        if str(nowMatch) in matchCount.keys():
            if matchCount[str(nowMatch)] >= 3:
                continue

        # 라운드가 끝나지 않았는데 배치될 경우 중단

        # 연전 기준에 맞지 않을 경우 중단
        # 휴일 연전 기준 체크
        if day > 0:
            flag = False

            if nowMatch in sch[day-1]:
                if day not in holidays:
                    continue
            
            # 전날이 휴일이 아닌 경우 
            if day not in holidays:
                for s in range(len(sch[day-1])):
                    if nowMatch[0] in sch[day-1][s] or nowMatch[1] in sch[day-1][s]:
                        flag = True
                        break

                if flag:
                    continue

            # 3연전 거르기 (전날 휴일인 경우)
            if day > 1 and day in holidays and day-1 in holidays:
                flag1 = False
                flag2 = False

                for s in range(len(sch[day-1])):
                    if nowMatch[0] in sch[day-1][s]:
                        flag1 = True
                        break
                
                for s in range(len(sch[day-2])):
                    if nowMatch[0] in sch[day-2][s]:
                        flag2 = True
                        break
                
                if flag1 and flag2:
                    continue

                flag1 = False
                flag2 = False

                for s in range(len(sch[day-1])):
                    if nowMatch[1] in sch[day-1][s]:
                        flag1 = True
                        break
                
                for s in range(len(sch[day-2])):
                    if nowMatch[1] in sch[day-2][s]:
                        flag2 = True
                        break
                
                if flag1 and flag2:
                    continue

        matchTmp.append(nowMatch)

        if len(matchTmp) == mustGameCount:
            tsch = sch[:]
            tsch.append(matchTmp)
            matchTmp = []

            # KPI 계산 하여 충족 되지 않을 경우 중단 (각 요소 별 최대-최소 2 초과 시)
            if KPI_func.calcKPI(tsch):
                q.append([day + 1, tsch])