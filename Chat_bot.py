import os
import pandas as pd

from enum import Enum
class USER_NOW_STEP(Enum):
    Not_Match = "with_chatbot"
    CHAT = "chat with person"
    Real = "in_person"


class SOCIAL_CHAT_BOT:
    def __init__(self, Result_File = "risk_analysis_result.xlsx"):
        self.result_file = Result_File
        self.df_risk = self.Load_Risk_Data()

        self.User_profile ={
            "user_ID" : "USER_CODE_001",
            "age_group" : "20S",
            "user_sex" : "남성",
            "region" : "서울 강남구",
            "user_intr" : [],
            "now_step" : USER_NOW_STEP.Not_Match
        }
        self.Matched_person = None

    def Load_Risk_Data(self):
        if os.path.exists(self.result_file):
            try:
                return pd.read_excel(self.result_file, sheet_name = "전체_위험도")
            except Exception:
                return None

        return None

    def User_Interests(self, message):
        Key_word = ["낚시", "사진", "등산", "산책", "음악", "요리", "동물", "독서", "영화", "운동"]
        return [intr for intr in Key_word if intr in message]
    
    def Process_Message(self, message):
        chat_intr = self.User_Interests(message)
        for k_word in chat_intr:
            if k_word not in self.User_profile["user_intr"]:
                self.User_profile["user_intr"].append(k_word)

        match self.User_profile["now_step"]:
            case USER_NOW_STEP.Not_Match:
                return "현재 진행 상태 : 매칭 상대 없음"
            case USER_NOW_STEP.CHAT:
                return "현재 진행 상태 : 채팅 상대 존재"
            case USER_NOW_STEP.Real:
                return "현재 진행 상태 : 현실 만남 성사"
            case _:
                return "현재 진행 상태 : 입력값 없음"

        
