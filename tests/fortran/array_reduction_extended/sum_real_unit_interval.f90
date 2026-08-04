! vybe-test: fortran/array_reduction_extended/sum_real_unit_interval
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
real :: a(4) = [0.5, 1.5, 2.5, 3.5]
if ((sum(a)) /= 8) then
    print *, "FAIL: want [8] got [", sum(a), "]"
    stop 1
end if
end program t
