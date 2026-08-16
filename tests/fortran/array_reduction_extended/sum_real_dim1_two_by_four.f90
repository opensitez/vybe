! vybe-test: fortran/array_reduction_extended/sum_real_dim1_two_by_four
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
real :: m(2,4) = reshape([1.0,2.0,3.0,4.0,5.0,6.0,7.0,8.0],[2,4])
real :: c(4)
c = sum(m, dim=1)
if ((c(1)) /= 3.00000000) then
    print *, "FAIL: want [3.00000000] got [", c(1), "]"
    stop 1
end if
if ((c(4)) /= 15.0000000) then
    print *, "FAIL: want [15.0000000] got [", c(4), "]"
    stop 1
end if
end program t
