! vybe-test: fortran/complex_extended/real_parts_not_equal
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a = (2.0, 0.0), b = (3.0, 0.0)
if ((merge(1, 0, real(a) /= real(b))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, real(a) /= real(b)), "]"
    stop 1
end if
end program t
