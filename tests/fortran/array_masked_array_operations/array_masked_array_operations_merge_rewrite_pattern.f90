! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_merge_rewrite_pattern
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_merge_rewrite_pattern
    integer :: values(4)
    integer :: result(4)
    values = (/ 1, 2, 3, 4 /)
    result = merge(values*3, values, values > 2)
    if ((sum(result)) /= 24) then
    print *, "FAIL: want [24] got [", sum(result), "]"
    stop 1
end if
    if ((result(2)) /= 2) then
    print *, "FAIL: want [2] got [", result(2), "]"
    stop 1
end if
    if ((result(3)) /= 9) then
    print *, "FAIL: want [9] got [", result(3), "]"
    stop 1
end if
end program array_masked_array_operations_merge_rewrite_pattern
