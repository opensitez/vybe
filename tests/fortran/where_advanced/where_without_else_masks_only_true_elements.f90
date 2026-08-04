! vybe-test: fortran/where_advanced/where_without_else_masks_only_true_elements
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: a(4) = [1, 2, 3, 4]
    where (a >= 3)
        a = a * 10
    end where
    if ((a(1)) /= 1) then
    print *, "FAIL: want [1] got [", a(1), "]"
    stop 1
end if
    if ((a(3)) /= 30) then
    print *, "FAIL: want [30] got [", a(3), "]"
    stop 1
end if
    if ((a(4)) /= 40) then
    print *, "FAIL: want [40] got [", a(4), "]"
    stop 1
end if
end program test
