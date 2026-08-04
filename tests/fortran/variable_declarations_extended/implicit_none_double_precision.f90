! vybe-test: fortran/variable_declarations_extended/implicit_none_double_precision
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
double precision :: d = 2.5d0
if (abs((d) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", d, "]"
    stop 1
end if
end program t
