! vybe-test: fortran/complex_extended/real_literal_9_2
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z = (9.0, 2.0)
if ((nint(real(z))) /= 9) then
    print *, "FAIL: want [9] got [", nint(real(z)), "]"
    stop 1
end if
end program t
