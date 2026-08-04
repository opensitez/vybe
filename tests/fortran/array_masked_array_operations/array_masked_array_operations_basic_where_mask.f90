! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_basic_where_mask
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_basic_where_mask
    integer, allocatable :: values(:)
    integer, allocatable :: result(:)
    values = (/ -1, 0, 1, 2, 3 /)
    result = values
    where (values >= 2)
        result = values * 10
    end where
    if ((result(1)) /= -1) then
    print *, "FAIL: want [-1] got [", result(1), "]"
    stop 1
end if
    if ((result(3)) /= 1) then
    print *, "FAIL: want [1] got [", result(3), "]"
    stop 1
end if
    if ((result(5)) /= 30) then
    print *, "FAIL: want [30] got [", result(5), "]"
    stop 1
end if
    if ((sum(result)) /= 34) then
    print *, "FAIL: want [34] got [", sum(result), "]"
    stop 1
end if
end program array_masked_array_operations_basic_where_mask
