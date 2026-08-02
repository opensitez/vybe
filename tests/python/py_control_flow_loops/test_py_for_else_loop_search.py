# vybe-test: python/py_control_flow_loops/test_py_for_else_loop_search
# origin: languages/python/tests/python/test_py_control_flow_loops.rs

def find_even(nums):
    for n in nums:
        if n % 2 == 0:
            print(f"Found even: {n}")
            break
    else:
        print("No even number found")

find_even([1, 3, 5, 6, 7])
find_even([1, 3, 5, 7])
