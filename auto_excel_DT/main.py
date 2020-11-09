import pandas as pd
import matplotlib.pyplot as plt

num = int(input('현재 데이터의 갯수가 몇개입니까? : '))
final_df = pd.DataFrame() # 데이터 프레임을 사용할 것이라는 명시만 한다
graph_seq = str(num) + '11' # 그래프의 순서 번호 / 2행 1열 1번째라고 읽으면 됨

for i in range(num):
    # 여기서는 데이터를 처리하는 구간입니다
    data_frame = pd.read_excel('./input/input' + str(i+1) + '.xlsx', sheet_name='개방결과정리') # 엑셀 파일 불러오기
    data_frame = data_frame[['① 광역시도내 기관별개방현황','Unnamed: 2','Unnamed: 3']][2:7] # top5 랭킹 데이터만 가져오기
    data_frame.columns = ['Ranking', 'Name', 'Quantity'] # 속성명이 이상하게 되어있으므로 다시 잡아주기

    # 여기서는 subplot 그래프의 데이터를 처리하는 구간입니다
    name_ls, amount_ls = [],[]
    for j in range(2, 7):
        name_ls.append(str(data_frame['Ranking'][j]) + '등(' + data_frame['Name'][j] + ')')
        amount_ls.append(data_frame['Quantity'][j])

    # 여기서는 그래프를 그리는 구간입니다
    plt.subplot(int(graph_seq))
    plt.rcParams["figure.figsize"] = (10, 3)
    plt.plot(name_ls, amount_ls, 'rs--')
    plt.rc('font', family='NanumGothic')
    plt.grid()

    for j, v in enumerate(name_ls):
        plt.text(v, amount_ls[j], amount_ls[j],  # 좌표 (x축 = v, y축 = y[0]..y[1], 표시 = y[0]..y[1])
                 fontsize=15,
                 color='black',
                 horizontalalignment='center',  # horizontalalignment (left, center, right)
                 verticalalignment='bottom')

    graph_seq = graph_seq[:2] + str(i + 2)  # 다음 그래프의 순서 번호를 지정

    data_frame = data_frame.append({'Ranking': '', 'Name': '', 'Quantity': ''}, ignore_index=True)  # 한칸 띄워주기
    final_df = pd.concat([final_df,data_frame])

# 세부 그래프간의 간격을 조정
plt.subplots_adjust(left=0.125,
                    bottom=0.1,
                    right=0.9,
                    top=0.9,
                    wspace=0.2,
                    hspace=0.35)

final_df.to_excel('result.xlsx',index = False)
plt.show()