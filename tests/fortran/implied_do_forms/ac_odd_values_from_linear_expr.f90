! vybe-test: fortran/implied_do_forms/ac_odd_values_from_linear_expr
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
integer :: a(4) = [(2 * i + 1, i = 1, 4)]
if ((sum(a)) /= 24) then
    print *, "FAIL: want [24] got [", sum(a), "]"
    stop 1
end if
end program t
