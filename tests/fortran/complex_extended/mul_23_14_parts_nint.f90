! vybe-test: fortran/complex_extended/mul_23_14_parts_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a, b, c
a = cmplx(2.0, 3.0)
b = cmplx(1.0, 4.0)
c = a * b
if ((nint(real(c))) /= -10) then
    print *, "FAIL: want [-10] got [", nint(real(c)), "]"
    stop 1
end if
if ((nint(aimag(c))) /= 11) then
    print *, "FAIL: want [11] got [", nint(aimag(c)), "]"
    stop 1
end if
end program t
