! vybe-test: fortran/complex_extended/cmplx_negative_real_positive_imag
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z
z = cmplx(-5.0, 3.0)
if ((nint(real(z))) /= -5) then
    print *, "FAIL: want [-5] got [", nint(real(z)), "]"
    stop 1
end if
if ((nint(aimag(z))) /= 3) then
    print *, "FAIL: want [3] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
