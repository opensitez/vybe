! vybe-test: fortran/complex_extended/conjg_twice_restores_parts
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z, c, r
z = cmplx(7.0, -2.0)
c = conjg(z)
r = conjg(c)
if ((nint(real(r))) /= 7) then
    print *, "FAIL: want [7] got [", nint(real(r)), "]"
    stop 1
end if
if ((nint(aimag(r))) /= -2) then
    print *, "FAIL: want [-2] got [", nint(aimag(r)), "]"
    stop 1
end if
end program t
