import os
import pandas as pd

class ANALYSIS_BOT:
    def __init__(self, Result_File = "risk_analysis_result.xlsx"):
        self.result_file = Result_File
        self.df_all = self.Load_File("전체_위험도")
        self.df_age = self.Load_File("연령대별_분석")
        self.df_act_chan = self.Load_File("활동_변화")
        self.df_act_div = self.Load_File("활동_다양성")
        self.df_act_conc = self.Load_File("활동_편증")

    def Load_File(self, file_label):
        if os.path.exists(self.result_file):
            try:
                return pd.read_excel(self.result_file, sheet_name = file_label)
            except Exception as e:
                print(f"[에러]'{file_label}' 불러오기 실패 : {e}")
                return None
        return None
    
    def Generate_Report(self):
        if self.df_all is None or self.df_age is None:
            return "[애러] 분석할 데이터가 없습니다."

        report_message = []
        report_message.append("=" * 65)
        report_message.append("[리포트] 사회적 고립험도 분석")
        report_message.append("=" * 65)

        report_message.append("\n[1] 연령대별 고립 위혐 요인")
        driver_summary = self.df_all.groupby(['AGE_COD', '주요기여항목']).size().unstack(fill_value=0)
        report_message.append(str(driver_summary))

        report_message.append("\n[2] 연령대별 고립 위험군 및 관심군 분포")
        high_risk_age = self.df_age[self.df_age['관심군_이상_비율'] > 0][['AGE_COD', '분석단위수', '평균_위험지수', '관심군_이상_비율']]
        report_message.append(high_risk_age.to_string(index=False))

        if 'MCT_SGG_CD' in self.df_all.columns:
            report_message.append("\n[3] 지역별 평균 위험 지수")
            region_summary = self.df_all.groupby('MCT_SGG_CD')['위험지수'].agg(['count', 'mean', 'max']).reset_index()
            region_summary.columns = ['지역', '세그먼트 수', '평균 위험지수', '최고 위험지수']
            report_message.append(region_summary.round(2).to.String(index=False))
