! vybe-test: fortran/implied_do_forms/ac_zero_origin_stride_four
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
integer :: a(3) = [(i, i = 0, 8, 4)]
if ((sum(a)) /= 12) then
    print *, "FAIL: want [12] got [", sum(a), "]"
    stop 1
end if
end program t
