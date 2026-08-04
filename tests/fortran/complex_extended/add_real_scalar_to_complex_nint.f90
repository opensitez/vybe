! vybe-test: fortran/complex_extended/add_real_scalar_to_complex_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z, r
z = cmplx(2.0, 3.0)
r = 1.0 + z
if ((nint(real(r))) /= 3) then
    print *, "FAIL: want [3] got [", nint(real(r)), "]"
    stop 1
end if
if ((nint(aimag(r))) /= 3) then
    print *, "FAIL: want [3] got [", nint(aimag(r)), "]"
    stop 1
end if
end program t
