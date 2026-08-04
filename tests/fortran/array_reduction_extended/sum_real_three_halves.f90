! vybe-test: fortran/array_reduction_extended/sum_real_three_halves
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
real :: a(3) = [1.5, 2.5, 3.5]
if (abs((sum(a)) - 7.5) > 1.0e-6) then
    print *, "FAIL: want [7.5] got [", sum(a), "]"
    stop 1
end if
end program t
