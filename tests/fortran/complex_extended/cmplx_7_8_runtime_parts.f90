! vybe-test: fortran/complex_extended/cmplx_7_8_runtime_parts
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z
z = cmplx(7.0, 8.0)
if ((nint(real(z))) /= 7) then
    print *, "FAIL: want [7] got [", nint(real(z)), "]"
    stop 1
end if
if ((nint(aimag(z))) /= 8) then
    print *, "FAIL: want [8] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
