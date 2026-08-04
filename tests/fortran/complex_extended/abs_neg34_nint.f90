! vybe-test: fortran/complex_extended/abs_neg34_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z
z = cmplx(-3.0, -4.0)
if ((nint(abs(z))) /= 5) then
    print *, "FAIL: want [5] got [", nint(abs(z)), "]"
    stop 1
end if
end program t
