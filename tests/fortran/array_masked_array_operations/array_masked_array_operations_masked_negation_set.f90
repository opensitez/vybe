! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_masked_negation_set
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_masked_negation_set
    integer :: values(5)
    integer :: result(5)
    values = (/ 3, 7, 11, 13, 17 /)
    where (values >= 10)
        result = 0
    else where
        result = values
    end where
    if ((sum(result)) /= 20) then
    print *, "FAIL: want [20] got [", sum(result), "]"
    stop 1
end if
    if ((result(1)) /= 3) then
    print *, "FAIL: want [3] got [", result(1), "]"
    stop 1
end if
    if ((result(3)) /= 0) then
    print *, "FAIL: want [0] got [", result(3), "]"
    stop 1
end if
    if ((result(5)) /= 0) then
    print *, "FAIL: want [0] got [", result(5), "]"
    stop 1
end if
end program array_masked_array_operations_masked_negation_set
