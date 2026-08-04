! vybe-test: fortran/variable_declarations_extended/double_complex_decl
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
double complex :: z = (3.0d0, 4.0d0)
if ((nint(real(z))) /= 3) then
    print *, "FAIL: want [3] got [", nint(real(z)), "]"
    stop 1
end if
if ((nint(aimag(z))) /= 4) then
    print *, "FAIL: want [4] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
