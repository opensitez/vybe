! vybe-test: fortran/random_number_extended/random_number_vector_norm_finite
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
real :: v(3)
call random_number(v)
if ((merge(1, 0, sqrt(v(1)**2 + v(2)**2 + v(3)**2) >= 0.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, sqrt(v(1)**2 + v(2)**2 + v(3)**2) >= 0.0), "]"
    stop 1
end if
end program t
