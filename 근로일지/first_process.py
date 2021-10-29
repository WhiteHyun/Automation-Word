from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from class_time import class_datetime_dict


def make_csv(year: int, month: int, file_name: str) -> None:
    with open(file_name, "w") as f:
        f.write("날짜,강의명,이미지,내용\n")

        start_date = datetime(year, month, 1)
        end_date = start_date + relativedelta(months=1) - timedelta(days=1)
        picture_index = 1
        for i in range(end_date.day):
            temp_date = start_date + timedelta(days=i)
            if (week := temp_date.weekday()) != 5 and week != 6:  # 주말은 제외
                for key, value in class_datetime_dict.items():
                    if week == value[0]:
                        f.write(
                            f"{temp_date.strftime('%m%d')},{key},{picture_index},\n"
                        )
                        picture_index += 1


if __name__ == "__main__":
    make_csv(2021, 10, "data.csv")
