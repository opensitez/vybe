! vybe-test: fortran/complex_extended/div_imag_over_imag_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a, b, c
a = cmplx(0.0, 8.0)
b = cmplx(0.0, 2.0)
c = a / b
if ((nint(real(c))) /= 4) then
    print *, "FAIL: want [4] got [", nint(real(c)), "]"
    stop 1
end if
if ((nint(aimag(c))) /= 0) then
    print *, "FAIL: want [0] got [", nint(aimag(c)), "]"
    stop 1
end if
end program t
