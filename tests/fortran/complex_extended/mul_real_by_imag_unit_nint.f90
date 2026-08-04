! vybe-test: fortran/complex_extended/mul_real_by_imag_unit_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a, b, c
a = cmplx(1.0, 0.0)
b = cmplx(0.0, 1.0)
c = a * b
if ((nint(real(c))) /= 0) then
    print *, "FAIL: want [0] got [", nint(real(c)), "]"
    stop 1
end if
if ((nint(aimag(c))) /= 1) then
    print *, "FAIL: want [1] got [", nint(aimag(c)), "]"
    stop 1
end if
end program t
