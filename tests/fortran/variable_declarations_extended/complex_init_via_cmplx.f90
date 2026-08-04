! vybe-test: fortran/variable_declarations_extended/complex_init_via_cmplx
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
complex :: z
z = cmplx(8.0, -1.0)
if ((nint(real(z))) /= 8) then
    print *, "FAIL: want [8] got [", nint(real(z)), "]"
    stop 1
end if
if ((nint(aimag(z))) /= -1) then
    print *, "FAIL: want [-1] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
