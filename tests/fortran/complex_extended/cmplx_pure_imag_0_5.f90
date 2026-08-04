! vybe-test: fortran/complex_extended/cmplx_pure_imag_0_5
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z
z = cmplx(0.0, 5.0)
if ((nint(real(z))) /= 0) then
    print *, "FAIL: want [0] got [", nint(real(z)), "]"
    stop 1
end if
if ((nint(aimag(z))) /= 5) then
    print *, "FAIL: want [5] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
