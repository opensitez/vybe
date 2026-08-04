! vybe-test: fortran/complex_extended/array_element_real_index_1
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: x(3)
x(1) = cmplx(11.0, 0.0)
x(2) = cmplx(22.0, 0.0)
x(3) = cmplx(33.0, 0.0)
if ((nint(real(x(1)))) /= 11) then
    print *, "FAIL: want [11] got [", nint(real(x(1))), "]"
    stop 1
end if
end program t
