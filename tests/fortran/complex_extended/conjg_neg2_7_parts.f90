! vybe-test: fortran/complex_extended/conjg_neg2_7_parts
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z, c
z = cmplx(-2.0, 7.0)
c = conjg(z)
if ((nint(real(c))) /= -2) then
    print *, "FAIL: want [-2] got [", nint(real(c)), "]"
    stop 1
end if
if ((nint(aimag(c))) /= -7) then
    print *, "FAIL: want [-7] got [", nint(aimag(c)), "]"
    stop 1
end if
end program t
