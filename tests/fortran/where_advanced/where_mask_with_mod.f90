! vybe-test: fortran/where_advanced/where_mask_with_mod
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: a(10) = [(i, i=1,10)]
    integer :: b(10)
    b = 0
    where (mod(a, 3) == 0)
        b = a
    end where
    if ((b(3)) /= 3) then
    print *, "FAIL: want [3] got [", b(3), "]"
    stop 1
end if
    if ((b(6)) /= 6) then
    print *, "FAIL: want [6] got [", b(6), "]"
    stop 1
end if
    if ((b(1)) /= 0) then
    print *, "FAIL: want [0] got [", b(1), "]"
    stop 1
end if
end program test
