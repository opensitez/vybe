! vybe-test: fortran/variable_declarations_extended/implicit_none_complex_print_real
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
complex :: z = (3.0, 4.0)
if ((nint(real(z))) /= 3) then
    print *, "FAIL: want [3] got [", nint(real(z)), "]"
    stop 1
end if
end program t
