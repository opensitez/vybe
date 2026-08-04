! vybe-test: fortran/complex_extended/conjg_5_neg3_parts
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z, c
z = cmplx(5.0, -3.0)
c = conjg(z)
if ((nint(real(c))) /= 5) then
    print *, "FAIL: want [5] got [", nint(real(c)), "]"
    stop 1
end if
if ((nint(aimag(c))) /= 3) then
    print *, "FAIL: want [3] got [", nint(aimag(c)), "]"
    stop 1
end if
end program t
