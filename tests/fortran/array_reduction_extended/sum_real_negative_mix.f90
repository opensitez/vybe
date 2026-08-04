! vybe-test: fortran/array_reduction_extended/sum_real_negative_mix
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
real :: a(4) = [2.0, -1.0, 3.0, -2.0]
if ((sum(a)) /= 2) then
    print *, "FAIL: want [2] got [", sum(a), "]"
    stop 1
end if
end program t
