! vybe-test: fortran/complex_extended/cmplx_kind8_parts_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
integer, parameter :: dp = kind(1.0d0)
complex(dp) :: z
z = cmplx(11.0_dp, 13.0_dp, dp)
if ((nint(real(z))) /= 11) then
    print *, "FAIL: want [11] got [", nint(real(z)), "]"
    stop 1
end if
if ((nint(aimag(z))) /= 13) then
    print *, "FAIL: want [13] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
