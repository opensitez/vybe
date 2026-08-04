! vybe-test: fortran/complex_extended/add_12_34_parts_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a, b, c
a = cmplx(1.0, 2.0)
b = cmplx(3.0, 4.0)
c = a + b
if ((nint(real(c))) /= 4) then
    print *, "FAIL: want [4] got [", nint(real(c)), "]"
    stop 1
end if
if ((nint(aimag(c))) /= 6) then
    print *, "FAIL: want [6] got [", nint(aimag(c)), "]"
    stop 1
end if
end program t
