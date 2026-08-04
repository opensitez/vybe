! vybe-test: fortran/implied_do_forms/ac_implied_do_with_parameter_bound
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
integer, parameter :: n = 4
integer :: a(n) = [(i, i = 1, n)]
if ((sum(a)) /= 10) then
    print *, "FAIL: want [10] got [", sum(a), "]"
    stop 1
end if
end program t
