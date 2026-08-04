! vybe-test: fortran/variable_declarations_extended/complex_literal_init_parts
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
complex :: z = (6.0, -2.0)
if ((nint(real(z))) /= 6) then
    print *, "FAIL: want [6] got [", nint(real(z)), "]"
    stop 1
end if
if ((nint(aimag(z))) /= -2) then
    print *, "FAIL: want [-2] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
