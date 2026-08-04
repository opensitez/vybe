! vybe-test: fortran/complex_extended/sub_real_scalar_from_complex_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z, r
z = cmplx(5.0, 7.0)
r = z - 2.0
if ((nint(real(r))) /= 3) then
    print *, "FAIL: want [3] got [", nint(real(r)), "]"
    stop 1
end if
if ((nint(aimag(r))) /= 7) then
    print *, "FAIL: want [7] got [", nint(aimag(r)), "]"
    stop 1
end if
end program t
