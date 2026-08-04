! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_masked_abs_transform
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_masked_abs_transform
    integer :: values(6)
    integer :: result(6)
    values = (/ -6, 5, -4, 3, -2, 1 /)
    where (values < 0)
        result = -values
    else where
        result = values
    end where
    if ((result(1)) /= 6) then
    print *, "FAIL: want [6] got [", result(1), "]"
    stop 1
end if
    if ((result(2)) /= 5) then
    print *, "FAIL: want [5] got [", result(2), "]"
    stop 1
end if
    if ((result(3)) /= 4) then
    print *, "FAIL: want [4] got [", result(3), "]"
    stop 1
end if
    if ((sum(result)) /= 21) then
    print *, "FAIL: want [21] got [", sum(result), "]"
    stop 1
end if
end program array_masked_array_operations_masked_abs_transform
