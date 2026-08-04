! vybe-test: fortran/implied_do_forms/ac_stride_three_corners
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
integer :: a(4) = [(i, i = 2, 11, 3)]
if ((a(1)) /= 2) then
    print *, "FAIL: want [2] got [", a(1), "]"
    stop 1
end if
if ((a(4)) /= 11) then
    print *, "FAIL: want [11] got [", a(4), "]"
    stop 1
end if
end program t
