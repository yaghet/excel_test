"""
Задача №1
Условие задачи:
Напишите алгоритм, который преобразует целое число в римскую цифру.
Римские цифры представлены семью различными символами:
I (1), V (5), X (10), L (50), C (100), D (500), M (1000).

Пример 1
Input: num = 3749
Output: "MMMDCCXLIX"

Пример 2
Input: num = 58
Output: "LVIII"

Пример 3
Input: num = 1994
Output: "MCMXCIV"
"""

romans = {
    1000: 'M',
    900: 'CM',
    500: 'D',
    400: 'CD',
    100: 'C',
    90: 'XC',
    50: 'L',
    40: 'XL',
    10: 'X',
    9: 'IX',
    5: 'V',
    4: 'IV',
    1: 'I',
}


def convert_integer_to_roman(numeral: int) -> str:
    result = list()
    for key, value in romans.items():
        while key <= numeral:
            result.append(value)
            numeral -= key

    return ''.join(result)


def test_func(cases: dict[int, str]) -> None:
    for case, expected_result in cases.items():
        result = convert_integer_to_roman(numeral=case)
        if result == expected_result:
            print(f'Test for the case {case} was successful!')
        else:
            print('Something went wrong')
            break

    else:
        print('\nAll tests were successful.')


if __name__ == '__main__':
    test_data = {
        3749: 'MMMDCCXLIX',
        58: 'LVIII',
        1994: 'MCMXCIV',
    }
    test_func(cases=test_data)
