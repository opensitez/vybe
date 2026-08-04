! vybe-test: fortran/complex_extended/real_part_greater_than
! origin: languages/fortran/tests/fortran/test_complex_extended.rs
program t
complex :: a = (7.0, 2.0), b = (2.0, 7.0)
if ((merge(1, 0, real(a) > real(b))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, real(a) > real(b)), "]"
    stop 1
end if
end program t
