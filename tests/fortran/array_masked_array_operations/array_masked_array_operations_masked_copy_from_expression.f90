! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_masked_copy_from_expression
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_masked_copy_from_expression
    integer :: values(6)
    integer :: result(6)
    values = (/ 1, 2, 3, 4, 5, 6 /)
    where (values >= 3)
        result = values + 10
    end where
    if ((sum(result)) /= 47) then
    print *, "FAIL: want [47] got [", sum(result), "]"
    stop 1
end if
    if ((result(1)) /= 0) then
    print *, "FAIL: want [0] got [", result(1), "]"
    stop 1
end if
    if ((result(4)) /= 14) then
    print *, "FAIL: want [14] got [", result(4), "]"
    stop 1
end if
    if ((result(6)) /= 16) then
    print *, "FAIL: want [16] got [", result(6), "]"
    stop 1
end if
end program array_masked_array_operations_masked_copy_from_expression
