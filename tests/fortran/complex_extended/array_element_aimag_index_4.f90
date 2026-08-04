! vybe-test: fortran/complex_extended/array_element_aimag_index_4
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: x(4)
x(1) = cmplx(0.0, 1.0)
x(2) = cmplx(0.0, 2.0)
x(3) = cmplx(0.0, 3.0)
x(4) = cmplx(0.0, 4.0)
if ((nint(aimag(x(4)))) /= 4) then
    print *, "FAIL: want [4] got [", nint(aimag(x(4))), "]"
    stop 1
end if
end program t
