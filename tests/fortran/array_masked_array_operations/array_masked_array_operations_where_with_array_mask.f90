! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_where_with_array_mask
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_where_with_array_mask
    integer :: values(5)
    integer :: mask(5)
    integer :: result(5)
    values = (/ 4, 5, 6, 7, 8 /)
    mask = (/ 1, 0, 1, 0, 1 /)
    where (mask == 1)
        result = values + 1
    end where
    if ((sum(result)) /= 21) then
    print *, "FAIL: want [21] got [", sum(result), "]"
    stop 1
end if
    if ((result(1)) /= 5) then
    print *, "FAIL: want [5] got [", result(1), "]"
    stop 1
end if
    if ((result(2)) /= 0) then
    print *, "FAIL: want [0] got [", result(2), "]"
    stop 1
end if
    if ((result(5)) /= 9) then
    print *, "FAIL: want [9] got [", result(5), "]"
    stop 1
end if
end program array_masked_array_operations_where_with_array_mask
