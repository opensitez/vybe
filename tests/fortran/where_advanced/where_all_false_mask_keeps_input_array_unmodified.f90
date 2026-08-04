! vybe-test: fortran/where_advanced/where_all_false_mask_keeps_input_array_unmodified
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: b(4)
    b = 0
    where (a < 0)
        b = 9
    end where
    if ((b(1)) /= 0) then
    print *, "FAIL: want [0] got [", b(1), "]"
    stop 1
end if
    if ((b(2)) /= 0) then
    print *, "FAIL: want [0] got [", b(2), "]"
    stop 1
end if
    if ((b(3)) /= 0) then
    print *, "FAIL: want [0] got [", b(3), "]"
    stop 1
end if
    if ((b(4)) /= 0) then
    print *, "FAIL: want [0] got [", b(4), "]"
    stop 1
end if
end program test
