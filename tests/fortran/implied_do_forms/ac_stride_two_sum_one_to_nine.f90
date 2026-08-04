! vybe-test: fortran/implied_do_forms/ac_stride_two_sum_one_to_nine
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
integer :: a(5) = [(i, i = 1, 9, 2)]
if ((sum(a)) /= 25) then
    print *, "FAIL: want [25] got [", sum(a), "]"
    stop 1
end if
end program t
