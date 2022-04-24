import base64
import datetime

import pandas as pd
import pytz
from icalendar import Calendar, Event


class csvtodf():
    def __init__(self, input_path):
        self.input_path=input_path
        self.df=pd.read_csv(input_path, skiprows=1)
        

    def df_input(self):
        """Original code: https://github.com/yu9824/member_decision main.py main function. 
        Partially changed for menber_decision_mip function.
        """
        index_Unnamed = [i for i, col in enumerate(self.df.columns) if 'Unnamed: ' in col]
        df_kagisime_input = self.df.iloc[:, index_Unnamed[0]+1:index_Unnamed[1]]
        df_gomisute_input = self.df.iloc[:, index_Unnamed[1]+1:index_Unnamed[2] if len(index_Unnamed) > 2 else None]

        #割り当て不可→False，割り当て可能→True
        df_gomisute_input.iloc[:, 2:]=df_gomisute_input.iloc[:, 2:].isnull()
        df_kagisime_input.iloc[:, 2:]=df_kagisime_input.iloc[:, 2:].isnull() 
        #csv読み込みの際のカラム名重複回避の.1を除去
        df_gomisute=df_gomisute_input.rename(columns=lambda s: s.strip('.1'))
        df_kagisime=df_kagisime_input.copy()

        df_gomisute['参加可能人数']=df_gomisute.iloc[:, 2:].sum(axis=1).astype(int)
        df_kagisime['参加可能人数']=df_kagisime.iloc[:, 2:].sum(axis=1).astype(int)
        df_gomisute['必要人数']=[4 if df_gomisute['参加可能人数'][i]>4 else df_gomisute['参加可能人数'][i] for i in df_gomisute.index.values]
        df_kagisime['必要人数']=[2 if df_kagisime['参加可能人数'][i]>2 else df_kagisime['参加可能人数'][i] for i in df_kagisime.index.values]

        self.df_kagisime=df_kagisime
        self.df_gomisute=df_gomisute
        self.N_kagisime_members=len(df_kagisime_input.columns)-2
        self.N_gomisute_members=len(df_gomisute_input.columns)-2
        yyyy, mm, _=self.df_kagisime['日付'][0].split('/')
        
        self.yyyymm=yyyy+mm

        return df_kagisime, df_gomisute

class makeical():
    def __init__(self, out_df, members=None):
        self.out_df=out_df
        yyyy, mm, _=self.out_df.index[0].split('/')
        
        self.yyyymm=yyyy+mm
        self.members=members
        if self.members is not None:
            self.files=self.convert_everyone()
    
    def convert(self, member, fname=None):
        member=member
        fname = member if fname is None else fname
            # カレンダーオブジェクトの生成
        cal = Calendar()

        # カレンダーに必須の項目
        cal.add('prodid', 'hito1280')
        cal.add('version', '2.0')

        # タイムゾーン
        tokyo = pytz.timezone('Asia/Tokyo')

        for name, series in self.out_df.iteritems():
            series_ = series[series.str.contains(member, na=False)]
            if name == '鍵閉め':
                start_td = datetime.timedelta(hours = 17, minutes = 45)   # 17時間45分
            elif name == 'ゴミ捨て':
                start_td = datetime.timedelta(hours = 12, minutes = 30)   # 12時間30分
            else:
                continue
            need_td = datetime.timedelta(hours = 1)

            for date, cell in zip(series_.index, series_):
                # 予定の開始時間と終了時間を変数として得る．
                start_time = datetime.datetime.strptime(date, '%Y/%m/%d') + start_td
                end_time = start_time + need_td

                # Eventオブジェクトの生成
                event = Event()

                # 必要情報
                event.add('summary', name)  # 予定名
                event.add('dtstart', tokyo.localize(start_time))
                event.add('dtend', tokyo.localize(end_time))
                event.add('description', cell)  # 誰とやるかを説明欄に記述
                event.add('created', tokyo.localize(datetime.datetime.now()))    # いつ作ったのか．

                # カレンダーに追加
                cal.add_component(event)

        # カレンダーのファイルへの書き出し
        file=cal.to_ical()
        encoded_file=base64.b64encode(cal.to_ical()).decode()
        filename=self.yyyymm+'_'+fname+'.ics'
        return encoded_file

    def convert_everyone(self):
        files=[self.convert(member) for member in self.members]
        return files


if __name__ == '__main__':
    # データの読み込み
    input_path = 'example/GSS_test - 202111 予定表.csv'
    data=csvtodf(input_path)
    print(data.df_input())

    print(data.N_gomisute_members)
