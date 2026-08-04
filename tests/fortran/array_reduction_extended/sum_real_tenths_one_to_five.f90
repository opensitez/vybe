! vybe-test: fortran/array_reduction_extended/sum_real_tenths_one_to_five
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
real :: a(5) = [0.1, 0.2, 0.3, 0.4, 0.5]
if (abs((sum(a)) - 1.5) > 1.0e-6) then
    print *, "FAIL: want [1.5] got [", sum(a), "]"
    stop 1
end if
end program t
