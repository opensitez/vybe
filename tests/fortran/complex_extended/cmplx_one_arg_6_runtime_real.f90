! vybe-test: fortran/complex_extended/cmplx_one_arg_6_runtime_real
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z
z = cmplx(6.0)
if ((nint(real(z))) /= 6) then
    print *, "FAIL: want [6] got [", nint(real(z)), "]"
    stop 1
end if
if ((nint(aimag(z))) /= 0) then
    print *, "FAIL: want [0] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
