! vybe-test: fortran/complex_extended/array_2d_element_real_part
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: m(2, 2)
m(1, 1) = cmplx(1.0, 2.0)
m(2, 1) = cmplx(5.0, 6.0)
if ((nint(real(m(2, 1)))) /= 5) then
    print *, "FAIL: want [5] got [", nint(real(m(2, 1))), "]"
    stop 1
end if
end program t
