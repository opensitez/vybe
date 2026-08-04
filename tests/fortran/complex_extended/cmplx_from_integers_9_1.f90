! vybe-test: fortran/complex_extended/cmplx_from_integers_9_1
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
integer :: i = 9, j = 1
complex :: z
z = cmplx(i, j)
if ((nint(real(z))) /= 9) then
    print *, "FAIL: want [9] got [", nint(real(z)), "]"
    stop 1
end if
if ((nint(aimag(z))) /= 1) then
    print *, "FAIL: want [1] got [", nint(aimag(z)), "]"
    stop 1
end if
end program t
