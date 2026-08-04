! vybe-test: fortran/complex_extended/aimag_of_negated_complex
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z, n
z = cmplx(4.0, -6.0)
n = -z
if ((nint(aimag(n))) /= 6) then
    print *, "FAIL: want [6] got [", nint(aimag(n)), "]"
    stop 1
end if
end program t
