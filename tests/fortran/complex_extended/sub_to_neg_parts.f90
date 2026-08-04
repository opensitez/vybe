! vybe-test: fortran/complex_extended/sub_to_neg_parts
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a, b, c
a = cmplx(0.0, 0.0)
b = cmplx(1.0, 1.0)
c = a - b
if ((nint(real(c))) /= -1) then
    print *, "FAIL: want [-1] got [", nint(real(c)), "]"
    stop 1
end if
if ((nint(aimag(c))) /= -1) then
    print *, "FAIL: want [-1] got [", nint(aimag(c)), "]"
    stop 1
end if
end program t
