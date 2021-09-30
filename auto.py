from docx import Document
from docx.shared import Cm

MONTH = 0
DAY = 1


def make_work_log(date: tuple, time: tuple, content: str, image_url: str, lecture: str):
    """

    Example
    -------
    >>> make_work_log(date=("9", "7"),
                      time=("0900", "1030"),
                      content="안녕하세요\\n만나서반갑습니다.",
                      image_url="test_image.jpg"
                      lecture="지능정보시스템")
    작성된 문서가 저장

    """
    doc = Document("양식.docx")
    tables = doc.tables

    work_day = tables[0].rows[3].cells[1].paragraphs[0]
    work_times = tables[0].rows[3].cells[3].paragraphs

    work_content = tables[0].rows[5].cells[1]
    work_picture = tables[0].rows[6].cells[1].paragraphs[0].add_run()

    # 근로일시 작성
    work_day.text = f"21년   {date[MONTH]}월  {date[DAY]}일"

    # 근로시간 작성
    for i, work_time in zip((0, 1), work_times):
        work_time.text = f"  {time[i][:2]}시  {time[i][2:]}분 {work_time.text[11:13]}"

    # 근로내용 작성
    for sentence in content.split("\n"):
        work_content.add_paragraph(sentence)

    # 근로사진 첨부

    work_picture.add_picture(image_url, width=Cm(11.62), height=Cm(7.24))

    doc.save(f"note/{date[MONTH]}월{date[DAY]}일 - {lecture}.docx")


if __name__ == "__main__":
    pass
