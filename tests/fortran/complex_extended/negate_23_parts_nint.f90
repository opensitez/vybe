! vybe-test: fortran/complex_extended/negate_23_parts_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: z, n
z = cmplx(2.0, 3.0)
n = -z
if ((nint(real(n))) /= -2) then
    print *, "FAIL: want [-2] got [", nint(real(n)), "]"
    stop 1
end if
if ((nint(aimag(n))) /= -3) then
    print *, "FAIL: want [-3] got [", nint(aimag(n)), "]"
    stop 1
end if
end program t
