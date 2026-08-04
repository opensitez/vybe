! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_masked_minmax_mix
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_masked_minmax_mix
    integer :: values(6)
    integer :: result(6)
    values = (/ 4, 1, 8, 2, 16, 3 /)
    result = -1
    where (values > 5)
        result = values
    end where
    if ((sum(result)) /= 20) then
    print *, "FAIL: want [20] got [", sum(result), "]"
    stop 1
end if
    if ((result(1)) /= -1) then
    print *, "FAIL: want [-1] got [", result(1), "]"
    stop 1
end if
    if ((result(3)) /= 8) then
    print *, "FAIL: want [8] got [", result(3), "]"
    stop 1
end if
    if ((result(6)) /= -1) then
    print *, "FAIL: want [-1] got [", result(6), "]"
    stop 1
end if
end program array_masked_array_operations_masked_minmax_mix
