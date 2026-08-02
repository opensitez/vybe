# vybe-test: python/programs/bubble_sort
# origin: languages/python/tests/python/test_programs.rs
# vybe-test-mode: compile

def bubble_sort(lst):
    n = len(lst)
    for i in range(n):
        for j in range(0, n - i - 1):
            if lst[j] > lst[j + 1]:
                lst[j], lst[j + 1] = lst[j + 1], lst[j]
    return lst

print(bubble_sort([64, 34, 25, 12, 22, 11, 90]))
