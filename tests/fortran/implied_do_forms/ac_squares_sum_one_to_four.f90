! vybe-test: fortran/implied_do_forms/ac_squares_sum_one_to_four
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
integer :: a(4) = [(i * i, i = 1, 4)]
if ((sum(a)) /= 30) then
    print *, "FAIL: want [30] got [", sum(a), "]"
    stop 1
end if
end program t
