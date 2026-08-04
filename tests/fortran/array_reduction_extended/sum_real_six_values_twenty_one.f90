! vybe-test: fortran/array_reduction_extended/sum_real_six_values_twenty_one
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
real :: a(6) = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0]
if ((sum(a)) /= 21) then
    print *, "FAIL: want [21] got [", sum(a), "]"
    stop 1
end if
end program t
