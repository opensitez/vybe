! vybe-test: fortran/complex_extended/array_elements_add_parts
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: x(2), y(2), z
x(1) = cmplx(1.0, 2.0)
x(2) = cmplx(3.0, 4.0)
y(1) = cmplx(5.0, 6.0)
y(2) = cmplx(7.0, 8.0)
z = x(1) + y(2)
if ((nint(real(z))) /= 8) then
    print *, "FAIL: want [8] got [", nint(real(z)), "]"
    stop 1
end if
if ((nint(aimag(z))) /= 10) then
    print *, "FAIL: want [10] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
