! vybe-test: fortran/implied_do_forms/ac_old_syntax_slash_paren_sum
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
integer :: a(4) = (/ (i, i = 1, 4) /)
if ((sum(a)) /= 10) then
    print *, "FAIL: want [10] got [", sum(a), "]"
    stop 1
end if
end program t
