! vybe-test: fortran/complex_extended/div_34_by_i_parts_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a, b, c
a = cmplx(3.0, 4.0)
b = cmplx(0.0, 1.0)
c = a / b
if ((nint(real(c))) /= 4) then
    print *, "FAIL: want [4] got [", nint(real(c)), "]"
    stop 1
end if
if ((nint(aimag(c))) /= -3) then
    print *, "FAIL: want [-3] got [", nint(aimag(c)), "]"
    stop 1
end if
end program t
