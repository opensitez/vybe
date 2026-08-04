! vybe-test: fortran/complex_extended/div_60_20_parts_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a, b, c
a = cmplx(6.0, 0.0)
b = cmplx(2.0, 0.0)
c = a / b
if ((nint(real(c))) /= 3) then
    print *, "FAIL: want [3] got [", nint(real(c)), "]"
    stop 1
end if
if ((nint(aimag(c))) /= 0) then
    print *, "FAIL: want [0] got [", nint(aimag(c)), "]"
    stop 1
end if
end program t
