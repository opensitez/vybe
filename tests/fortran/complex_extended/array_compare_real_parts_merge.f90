! vybe-test: fortran/complex_extended/array_compare_real_parts_merge
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: x(2)
x(1) = cmplx(4.0, 1.0)
x(2) = cmplx(4.0, 9.0)
if ((merge(1, 0, real(x(1)) == real(x(2)))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, real(x(1)) == real(x(2))), "]"
    stop 1
end if
end program t
