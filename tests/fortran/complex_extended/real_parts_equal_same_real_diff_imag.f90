! vybe-test: fortran/complex_extended/real_parts_equal_same_real_diff_imag
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a = (2.0, 9.0), b = (2.0, 1.0)
if ((merge(1, 0, real(a) == real(b))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, real(a) == real(b)), "]"
    stop 1
end if
end program t
