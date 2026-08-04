! vybe-test: fortran/complex_extended/array_element_abs_nint
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: x(2)
x(1) = cmplx(5.0, 12.0)
x(2) = cmplx(1.0, 1.0)
if ((nint(abs(x(1)))) /= 13) then
    print *, "FAIL: want [13] got [", nint(abs(x(1))), "]"
    stop 1
end if
end program t
