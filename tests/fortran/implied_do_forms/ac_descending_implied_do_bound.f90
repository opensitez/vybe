! vybe-test: fortran/implied_do_forms/ac_descending_implied_do_bound
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
integer, parameter :: n = 6
integer :: a(4) = [(i, i = n, 1, -2)]
if ((a(1)) /= 6) then
    print *, "FAIL: want [6] got [", a(1), "]"
    stop 1
end if
if ((a(4)) /= 2) then
    print *, "FAIL: want [2] got [", a(4), "]"
    stop 1
end if
end program t
