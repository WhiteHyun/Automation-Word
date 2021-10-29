import pandas as pd
from class_time import class_datetime_dict
import data


if __name__ == "__main__":
    df = pd.read_csv("data.csv", dtype={"날짜": str, "강의명": str, "이미지": str, "내용": str})

    # csv의 과목명과 global_time (dict)의 길이가 무조건 같아야 함
    assert len(set(df.loc[:, "강의명"])) == len(class_datetime_dict)
    for i in df.index:
        date, time, content, image_url = data.preprocessing(df.iloc[i, :])
        data.DailyLog(
            date=date,
            time=time,
            content=content,
            image_link=image_url,
            lecture=df.loc[i, "강의명"],
        ).make_work_log()
