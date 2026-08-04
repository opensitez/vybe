! vybe-test: fortran/complex_extended/conjg_of_added_complex
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a, b, s, c
a = cmplx(1.0, 2.0)
b = cmplx(3.0, 4.0)
s = a + b
c = conjg(s)
if ((nint(real(c))) /= 4) then
    print *, "FAIL: want [4] got [", nint(real(c)), "]"
    stop 1
end if
if ((nint(aimag(c))) /= -6) then
    print *, "FAIL: want [-6] got [", nint(aimag(c)), "]"
    stop 1
end if
end program t
