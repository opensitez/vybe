! vybe-test: fortran/complex_extended/real_of_conjg_part
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z, c
z = cmplx(8.0, -3.0)
c = conjg(z)
if ((nint(real(c))) /= 8) then
    print *, "FAIL: want [8] got [", nint(real(c)), "]"
    stop 1
end if
end program t
