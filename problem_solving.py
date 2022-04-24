import pandas as pd
from mip import Model, minimize, xsum

from read_df import csvtodf


class KandGProblem:
    """グループ分け問題を解くクラス"""

    def __init__(self, df_kagisime, df_gomisute, name="KandGProblem"):
        # 初期化メソッド
        self.df_kagisime = df_kagisime
        self.df_gomisute = df_gomisute
        self.name = name
        self.prob = self._formulate()

    def _formulate(self):
        # グループ分け問題（0-1整数計画問題）のインスタンス作成
        m = Model()

        # Constant
        N_days = self.df_gomisute.shape[0]
        L_gactivedays = [i for i, v in enumerate(self.df_gomisute["参加可能人数"]) if v != 0]
        L_kactivedays = [i for i, v in enumerate(self.df_kagisime["参加可能人数"]) if v != 0]
        N_gactivedays = len(L_gactivedays)
        N_kactivedays = len(L_kactivedays)
        N_gomisute_members = self.df_gomisute.shape[1] - 4
        N_kagisime_members = self.df_kagisime.shape[1] - 4
        L_ksplited = [
            L_kactivedays[idx : idx + 5] for idx in range(0, N_kactivedays, 5)
        ]
        N_weeks = len(L_ksplited)

        # Variable for optimization.
        V_kshift_table = m.add_var_tensor(
            (N_days, N_kagisime_members),
            name="V_kshift_table",
            var_type="INTEGER",
            lb=0,
            ub=1,
        )
        z_kequal_person = m.add_var_tensor(
            (N_kagisime_members,), name="z_kequal_person", var_type="INTEGER"
        )
        z_kequal_week = m.add_var_tensor(
            (N_weeks, N_kagisime_members), name="z_kequal_week", var_type="INTEGER"
        )

        V_gshift_table = m.add_var_tensor(
            (N_days, N_gomisute_members),
            name="V_gshift_table",
            var_type="INTEGER",
            lb=0,
            ub=1,
        )
        z_gequal_person = m.add_var_tensor(
            (N_gomisute_members,), name="z_gequal_person", var_type="CONTINUOUS"
        )

        z_sameday_kg = m.add_var_tensor(
            (N_gactivedays, N_kagisime_members),
            name="z_kequal_person",
            var_type="INTEGER",
        )

        C_equal_person = 1000
        Cl_equal_week = [100 for i in range(N_kagisime_members)]
        Cl_sameday_kg = [10 for i in range(N_kagisime_members)]

        # 目的関数
        m.objective = minimize(
            C_equal_person * xsum(z_kequal_person)
            + C_equal_person * xsum(z_gequal_person)
            + xsum(
                Cl_equal_week[i] * xsum(z_kequal_week[:, i])
                for i in range(N_kagisime_members)
            )
            + xsum(
                Cl_sameday_kg[i] * xsum(z_sameday_kg[:, i])
                for i in range(N_kagisime_members)
            )
        )

        # 制約条件：
        for i, r in self.df_kagisime.iloc[:, 2:].iterrows():
            # 入れない日には入らない（入れない日=0(False)）, 必要人数を満たす
            for j in range(N_kagisime_members):
                m += V_kshift_table[i][j] <= r[j]
            m += xsum(V_kshift_table[i]) == r["必要人数"]

        for i, r in self.df_gomisute.iloc[:, 2:].iterrows():
            for j in range(N_gomisute_members):
                m += V_gshift_table[i][j] <= r[j]
            m += xsum(V_gshift_table[i]) == r["必要人数"]
        # 絶対値に変換するための制限
        for i in range(N_kagisime_members):
            m += (
                xsum(V_kshift_table[:, i]) - (N_kactivedays * 2) // N_kagisime_members
            ) >= -z_kequal_person[i]
            m += (
                xsum(V_kshift_table[:, i]) - (N_kactivedays * 2) // N_kagisime_members
            ) <= z_kequal_person[i]
            for j, l_weekday in enumerate(L_ksplited):
                m += (
                    xsum(V_kshift_table[l_weekday, i])
                    - (len(l_weekday) * 2) // N_kagisime_members
                ) >= -z_kequal_week[j, i]
                m += (
                    xsum(V_kshift_table[l_weekday, i])
                    - (len(l_weekday) * 2) // N_kagisime_members
                ) <= z_kequal_week[j, i]
            # 差をとって絶対値にする->最小化：(k, g)->z (1, 0), (0, 1)->1, (1, 1), (0, 0)->0, (2, 0)->2, (2, 1)->1
            for j, v in enumerate(L_gactivedays):
                m += (V_kshift_table[v, i] - V_gshift_table[v, i]) >= -z_sameday_kg[
                    j, i
                ]
                m += (V_kshift_table[v, i] - V_gshift_table[v, i]) <= z_sameday_kg[j, i]
        for i in range(N_gomisute_members):
            m += (
                xsum(V_gshift_table[:, i]) - (N_gactivedays * 4) // N_gomisute_members
            ) >= -z_gequal_person[i]
            m += (
                xsum(V_gshift_table[:, i]) - (N_gactivedays * 4) // N_gomisute_members
            ) <= z_gequal_person[i]
            m += (
                xsum(V_gshift_table[:, i]) - (N_gactivedays * 4) // N_gomisute_members
            ) <= 1
            m += (
                xsum(V_gshift_table[:, i]) - (N_gactivedays * 4) // N_gomisute_members
            ) >= -1

        # 最適化後に利用するデータを返却
        return {
            "prob": m,
            "variable": {
                "V_kshift_table": V_kshift_table,
                "V_gshift_table": V_gshift_table,
            },
            "Constant": {
                "N_kagisime_members": N_kagisime_members,
                "N_gomisute_members": N_gomisute_members,
            },
        }

    def solve(self) -> pd.DataFrame:
        # 最適化問題を解くメソッド
        # 問題を解く
        self.prob["prob"].optimize()

        kagisime_shift_table = (
            self.prob["variable"]["V_kshift_table"].astype(float).astype(int)
        )
        gomisute_shift_table = (
            self.prob["variable"]["V_gshift_table"].astype(float).astype(int)
        )
        N_kagisime_members = self.prob["Constant"]["N_kagisime_members"]
        N_gomisute_members = self.prob["Constant"]["N_gomisute_members"]

        self.df_kagisime["Result"] = [
            ", ".join(
                j
                for i, j in zip(
                    r, self.df_kagisime.iloc[:, 2 : 2 + N_kagisime_members].columns
                )
                if i == 1
            )
            for r in kagisime_shift_table
        ]
        self.df_gomisute["Result"] = [
            ", ".join(
                j
                for i, j in zip(
                    r, self.df_gomisute.iloc[:, 2 : 2 + N_gomisute_members].columns
                )
                if i == 1
            )
            for r in gomisute_shift_table
        ]
        print("目的関数", self.prob["prob"].objective_value)
        print(self.df_kagisime[["日付", "曜日", "Result"]])
        print(self.df_gomisute[["日付", "曜日", "Result"]])

        df_output = pd.DataFrame(
            {
                "鍵閉め": list(self.df_kagisime["Result"]),
                "ゴミ捨て": list(self.df_gomisute["Result"]),
            },
            index=list(self.df_kagisime["日付"]),
        )
        L_gomisute_members = self.df_gomisute.iloc[
            :, 2 : 2 + N_gomisute_members
        ].columns

        self.L_gomisute_members = L_gomisute_members

        return df_output


if __name__ == "__main__":
    # データの読み込み
    input_path = "example/GSS_test - 202111 予定表.csv"
    data = csvtodf(input_path)
    df_kagisime, df_gomisute = data.df_input()

    # 数理モデル インスタンスの作成
    prob = KandGProblem(df_kagisime, df_gomisute)

    # 問題を解く
    solution_df = prob.solve()

    # 結果の表示
    print("Solution: \n", solution_df)
