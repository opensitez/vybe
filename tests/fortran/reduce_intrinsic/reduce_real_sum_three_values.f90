! vybe-test: fortran/reduce_intrinsic/reduce_real_sum_three_values
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
real :: vals(3) = [1.5, 2.5, 3.0]
if ((reduce(vals, operator(+))) /= 7) then
    print *, "FAIL: want [7] got [", reduce(vals, operator(+)), "]"
    stop 1
end if
end program t
