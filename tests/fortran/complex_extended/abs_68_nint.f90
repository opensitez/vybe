! vybe-test: fortran/complex_extended/abs_68_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z
z = cmplx(6.0, 8.0)
if ((nint(abs(z))) /= 10) then
    print *, "FAIL: want [10] got [", nint(abs(z)), "]"
    stop 1
end if
end program t
