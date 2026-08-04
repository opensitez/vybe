! vybe-test: fortran/complex_extended/cmplx_expression_args
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z
z = cmplx(1.0 + 4.0, 2.0 + 2.0)
if ((nint(real(z))) /= 5) then
    print *, "FAIL: want [5] got [", nint(real(z)), "]"
    stop 1
end if
if ((nint(aimag(z))) /= 4) then
    print *, "FAIL: want [4] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
