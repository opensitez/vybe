! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_categorical_masked_counts
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_categorical_masked_counts
    integer :: values(8)
    integer :: cat_a
    integer :: cat_b
    values = (/ 1, 2, 3, 4, 5, 6, 7, 8 /)
    cat_a = count(values <= 4)
    cat_b = count(values > 4)
    if ((cat_a) /= 4) then
    print *, "FAIL: want [4] got [", cat_a, "]"
    stop 1
end if
    if ((cat_b) /= 4) then
    print *, "FAIL: want [4] got [", cat_b, "]"
    stop 1
end if
    if ((sum(merge(1, 0, values < 3))) /= 2) then
    print *, "FAIL: want [2] got [", sum(merge(1, 0, values < 3)), "]"
    stop 1
end if
end program array_masked_array_operations_categorical_masked_counts
