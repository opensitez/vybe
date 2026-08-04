! vybe-test: fortran/complex_extended/abs_z_plus_conjg_half_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z, s
z = cmplx(3.0, 4.0)
s = z + conjg(z)
if ((nint(abs(s) / 2.0)) /= 3) then
    print *, "FAIL: want [3] got [", nint(abs(s) / 2.0), "]"
    stop 1
end if
end program t
