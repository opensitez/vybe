! vybe-test: fortran/complex_extended/abs_512_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z
z = cmplx(5.0, 12.0)
if ((nint(abs(z))) /= 13) then
    print *, "FAIL: want [13] got [", nint(abs(z)), "]"
    stop 1
end if
end program t
