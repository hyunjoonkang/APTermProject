'''
Name: Convert_Program
Version: 1.0.1
Summary: 코드 맨 밑줄에 main 함수를 선언해 두었습니다. 

 *아래 코드에서 반드시 사용환경에 맞는 파일주소를 입력해주세요!
'''

# import
import openpyxl

# function definition


def h2r(text, target_idx):  # 한글-로마자 함수
    # 리스트 생성
    start_list = ["ㄱ", "ㄲ", "ㄴ", "ㄷ", "ㄸ", "ㄹ", "ㅁ", "ㅂ", "ㅃ",
                  "ㅅ", "ㅆ", "ㅇ", "ㅈ", "ㅉ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ"]  # 19개
    middle_list = ["ㅏ", "ㅐ", "ㅑ", "ㅒ", "ㅓ", "ㅔ", "ㅕ", "ㅖ", "ㅗ", "ㅘ",
                   "ㅙ", "ㅚ", "ㅛ", "ㅜ", "ㅝ", "ㅞ", "ㅟ", "ㅠ", "ㅡ", "ㅢ", "ㅣ"]  # 21개
    end_list = ["", "ㄱ", "ㄲ", "ㄳ", "ㄴ", "ㄵ", "ㄶ", "ㄷ", "ㄹ", "ㄺ", "ㄻ", "ㄼ", "ㄽ", "ㄾ",
                "ㄿ", "ㅀ", "ㅁ", "ㅂ", "ㅄ", "ㅅ", "ㅆ", "ㅇ", "ㅈ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ"]  # 28개

    # 매칭 딕셔너리 생성
    start_dict = excel_to_dict(1, target_idx)
    start_dict.update({"ㅇ": ""})  # 예외: 엑셀에서 "" 저장 불가한 관계로 프로그램 내에서 직접 업데이트함

    middle_dict = excel_to_dict(2, target_idx)

    end_dict = excel_to_dict(3, target_idx)

    result = ""  # 변환된 언어가 저장될 string변수

    # 변환알고리즘
    for char in text:
        if "가" <= char <= "힣":
            # 한글 유니코드 계산법을 역추적하여 초성, 중성, 종성으로 분해하여 저장
            char_code = ord(char) - ord("가")
            start = start_list[char_code // (21 * 28)]
            middle = middle_list[(char_code % (21 * 28)) // 28]
            end = end_list[char_code % 28]

            # 딕셔너리 매칭함수
            result += start_dict.get(start, start) + \
                middle_dict.get(middle, middle) + end_dict.get(end, end)
        else:
            result += char  # 딕셔너리에 매치되지 않는다면 변환하지 않고 그대로 저장

    return result  # string 반환


def train(text, target_idx):  # 한글 -> 선택언어 변환 함수

    train_dic = excel_to_dict(0, target_idx)

    return train_dic.get(text, "해당하는 역명을 찾을 수 없습니다.")


def excel_to_dict(sheet_idx, target_idx):  # 엑셀파일을 입력받아 딕셔너리 생성하는 함수
    wb = openpyxl.load_workbook("data_Ver.1.0.xlsx")  # 사용환경에 맞는 주소를 입력해주세요.

    sheet = wb.worksheets[sheet_idx]

    data_dict = {}

    # 각 행을 반복하면서 데이터 읽기
    # 첫 번째 행은 헤더이므로 무시하고 2행부터 시작
    for row in sheet.iter_rows(min_row=1, values_only=True):
        key = row[0]  # 첫 번째 열의 값
        value = row[target_idx]  # 두 번째 열의 값
        data_dict[key] = value

    # 엑셀 파일 닫기
    wb.close()

    return data_dict


# main function
def main():
    start_idx = 0
    # 시작 화면
    print("\n발음 변환 프로그램 1.0")
    print("Made by 천홍규, 강현준, 최수정, 김예서, 이승환 | 가천대학교 소프트웨어학과")
    print("\n본 프로그램은 '한글 발음'을 '다른 언어의 발음'으로 변환하는 프로그램입니다.\n\n다음과 같은 기능을 포함합니다.\n\n----------------------")
    print("1 = 지하철역 변환")
    print("2 = 한글 - 외래어 변환")
    input("----------------------\n\n모두 읽으셨다면 엔터를 눌러주세요.")

    mode = int(input("\n사용할 기능의 번호를 입력해주세요: "))
    # 지하철 번역모드
    if mode == 1:
        print("\n\n----------------------\n 1 = 영어\n 2 = 일본어\n 3 = 중국어\n 4 = 러시아어\n 5 = 태국어\n----------------------\n")
        target_idx = int(input("변환하고 싶은 언어의 번호를 적으세요: "))
        text = str(input("변환하고 싶은 지하철역을 입력하세요\n역명에서 '역'자는 제외하고 적으세요! (ex 가천대): "))
        changed = train(text, target_idx)
        print(f"\n{text} \n=> {changed}\n")
    # 외래어 번역모드
    elif mode == 2:
        print("\n----------------------\n 1 = 영어\n *(준비중) = 프랑스어\n *(준비중) = 러시아어\n----------------------\n")
        target_idx = int(input("변환하고 싶은 언어의 번호를 적으세요: "))
        if target_idx == 1:
            text = str(input("변환하고 싶은 한글 문자열을 입력하세요: "))
            change = h2r(text, target_idx)
            print(f"\n{text} \n=> {change}\n")
        else:
            print("유효하지 않은 값입니다.")
    # 예외 처리
    else:
        print("유효하지 않은 값입니다.")


# main function call (starting program)
if __name__ == '__main__':
    main()
