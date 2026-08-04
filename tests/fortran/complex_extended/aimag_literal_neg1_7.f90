! vybe-test: fortran/complex_extended/aimag_literal_neg1_7
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z = (-1.0, 7.0)
if ((nint(aimag(z))) /= 7) then
    print *, "FAIL: want [7] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
