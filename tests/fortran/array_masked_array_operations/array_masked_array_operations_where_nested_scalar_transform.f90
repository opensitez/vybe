! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_where_nested_scalar_transform
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_where_nested_scalar_transform
    integer :: values(4)
    integer :: result(4)
    values = (/ 2, 3, 4, 5 /)
    where (mod(values,2) == 0)
        where (values > 3)
            result = values * 3
        elsewhere
            result = values * 2
        end where
    end where
    if ((result(1)) /= 4) then
    print *, "FAIL: want [4] got [", result(1), "]"
    stop 1
end if
    if ((result(2)) /= 6) then
    print *, "FAIL: want [6] got [", result(2), "]"
    stop 1
end if
    if ((result(3)) /= 12) then
    print *, "FAIL: want [12] got [", result(3), "]"
    stop 1
end if
    if ((result(4)) /= 15) then
    print *, "FAIL: want [15] got [", result(4), "]"
    stop 1
end if
    if ((sum(result)) /= 37) then
    print *, "FAIL: want [37] got [", sum(result), "]"
    stop 1
end if
end program array_masked_array_operations_where_nested_scalar_transform
