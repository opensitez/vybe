! vybe-test: fortran/integer_mod_division/integer_division_with_variables_truncates
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
integer :: a = -29, b = 6
if ((a / b) /= -4) then
    print *, "FAIL: want [-4] got [", a / b, "]"
    stop 1
end if
end program t
