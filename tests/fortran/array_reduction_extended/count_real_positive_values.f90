! vybe-test: fortran/array_reduction_extended/count_real_positive_values
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
real :: a(5) = [-1.0, 0.0, 1.5, -2.0, 3.0]
if ((count(a > 0.0)) /= 2) then
    print *, "FAIL: want [2] got [", count(a > 0.0), "]"
    stop 1
end if
end program t
