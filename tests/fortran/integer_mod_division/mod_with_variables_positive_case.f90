! vybe-test: fortran/integer_mod_division/mod_with_variables_positive_case
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
integer :: a = 29, b = 6
if ((mod(a, b)) /= 5) then
    print *, "FAIL: want [5] got [", mod(a, b), "]"
    stop 1
end if
end program t
