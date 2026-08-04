! vybe-test: fortran/array_masked_array_operations/array_masked_array_operations_where_elsewhere
! origin: languages/fortran/tests/fortran/test_array_masked_array_operations.rs

program array_masked_array_operations_where_elsewhere
    integer :: values(6)
    integer :: replaced(6)
    values = (/ 1, -2, 3, -4, 5, 0 /)
    where (values > 0)
        replaced = 100
    elsewhere
        replaced = -100
    end where
    if ((replaced(2)) /= -100) then
    print *, "FAIL: want [-100] got [", replaced(2), "]"
    stop 1
end if
    if ((replaced(3)) /= 100) then
    print *, "FAIL: want [100] got [", replaced(3), "]"
    stop 1
end if
    if ((replaced(6)) /= -100) then
    print *, "FAIL: want [-100] got [", replaced(6), "]"
    stop 1
end if
    if ((sum(replaced)) /= 0) then
    print *, "FAIL: want [0] got [", sum(replaced), "]"
    stop 1
end if
end program array_masked_array_operations_where_elsewhere
