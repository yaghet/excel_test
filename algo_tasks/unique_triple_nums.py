"""
Задача №2
Условие задачи:
Напишите функцию, которая принимает массив чисел nums и возвращает все уникальные тройки чисел [nums[i], nums[j], nums[k]], такие что:

i != j, i != k, j != k (индексы разные);
nums[i] + nums[j] + nums[k] == 0 (сумма тройки равна нулю).

Правила:
В решении не должно быть повторяющихся троек (даже если исходный массив содержит дубликаты).
Порядок троек и чисел внутри них не важен.

Пример 1
Input: nums = [-1, 0, 1]
Output: [[-1, 0, 1]]

Пример 2
Input: nums = [0, 0, 0, 0]
Output: [[0, 0, 0]]

Пример 3
Input: nums = [-2, 0, 1, 1, 2]
Output: [[-2, 0, 2], [-2, 1, 1]]
"""


def find_all_unique_triple_nums(array: list[int]) -> list[list[int]]:
    array.sort()  # Time: O(n log n), Space (1)
    n = len(array)  # Space: O(1)
    result = list()  # Space: O(k)

    for i in range(n - 2): # Time: O(n)
        if i > 0 and array[i] == array[i - 1]:
            continue

        left, right = i + 1, n - 1

        while left < right: # Time: O(n)
            cur_result = [array[i], array[left], array[right]]

            if sum(cur_result) == 0:
                result.append(cur_result)

                while left < right and array[left] == array[left + 1]:
                    left += 1

                while left < right and array[right] == array[right - 1]:
                    right -= 1

                left += 1
                right -= 1

            elif sum(cur_result) > 0:
                right -= 1
            else:
                left += 1

    # Total Time complexity: O(n ** 2), Space complexity: O(k)

    return result


def test_cases(cases: dict[tuple[int], list[int]]):
    for case, expected_result in cases.items():
        result = find_all_unique_triple_nums(array=list(case))
        if result == expected_result:
            print(f'Test for case {case} was successful')
        else:
            print('Something went wrong')
            break

    else:
        print('\nAll tests were successful')


if __name__ == '__main__':
    test_data = {
        (-1, 0, 1): [[-1, 0, 1]],
        (0, 0, 0, 0): [[0, 0, 0]],
        (-2, 0, 1, 1, 2): [[-2, 0, 2], [-2, 1, 1]]
    }

    test_cases(cases=test_data)
