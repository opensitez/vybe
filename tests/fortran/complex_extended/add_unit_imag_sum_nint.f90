! vybe-test: fortran/complex_extended/add_unit_imag_sum_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a, b, c
a = cmplx(0.0, 1.0)
b = cmplx(0.0, 1.0)
c = a + b
if ((nint(real(c))) /= 0) then
    print *, "FAIL: want [0] got [", nint(real(c)), "]"
    stop 1
end if
if ((nint(aimag(c))) /= 2) then
    print *, "FAIL: want [2] got [", nint(aimag(c)), "]"
    stop 1
end if
end program t
