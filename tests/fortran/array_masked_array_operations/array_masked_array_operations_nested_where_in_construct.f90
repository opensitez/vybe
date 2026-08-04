! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_nested_where_in_construct
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_nested_where_in_construct
    integer :: values(5)
    integer :: result(5)
    values = (/ 12, 11, 10, 9, 8 /)
    if (all(values > 0)) then
        where (mod(values, 2) == 0)
            result = 1
        elsewhere
            result = 0
        end where
    else
        result = -1
    end if
    if ((sum(result)) /= 3) then
    print *, "FAIL: want [3] got [", sum(result), "]"
    stop 1
end if
    if ((result(1)) /= 1) then
    print *, "FAIL: want [1] got [", result(1), "]"
    stop 1
end if
    if ((result(2)) /= 0) then
    print *, "FAIL: want [0] got [", result(2), "]"
    stop 1
end if
end program array_masked_array_operations_nested_where_in_construct
