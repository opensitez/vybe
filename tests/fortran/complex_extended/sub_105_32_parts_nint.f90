! vybe-test: fortran/complex_extended/sub_105_32_parts_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a, b, c
a = cmplx(10.0, 5.0)
b = cmplx(3.0, 2.0)
c = a - b
if ((nint(real(c))) /= 7) then
    print *, "FAIL: want [7] got [", nint(real(c)), "]"
    stop 1
end if
if ((nint(aimag(c))) /= 3) then
    print *, "FAIL: want [3] got [", nint(aimag(c)), "]"
    stop 1
end if
end program t
