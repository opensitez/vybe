! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_masked_elementary_sum
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_masked_elementary_sum
    integer :: values(5)
    integer :: masked_sum
    values = (/ 10, 20, 5, 15, 8 /)
    masked_sum = sum(values, values > 9)
    if ((masked_sum) /= 45) then
    print *, "FAIL: want [45] got [", masked_sum, "]"
    stop 1
end if
    if ((count(values > 9)) /= 3) then
    print *, "FAIL: want [3] got [", count(values > 9), "]"
    stop 1
end if
end program array_masked_array_operations_masked_elementary_sum
