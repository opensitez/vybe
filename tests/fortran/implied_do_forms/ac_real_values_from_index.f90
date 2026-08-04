! vybe-test: fortran/implied_do_forms/ac_real_values_from_index
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
real :: r(3) = [(real(i) + 0.5, i = 1, 3)]
if (abs((r(2)) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", r(2), "]"
    stop 1
end if
end program t
