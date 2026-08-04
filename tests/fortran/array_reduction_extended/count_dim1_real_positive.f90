! vybe-test: fortran/array_reduction_extended/count_dim1_real_positive
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
real :: m(2,3) = reshape([1.0,-1.0,2.0,3.0,-2.0,4.0],[2,3])
integer :: c(3)
c = count(m > 0.0, dim=1)
if ((c(1)) /= 2) then
    print *, "FAIL: want [2] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= 1) then
    print *, "FAIL: want [1] got [", c(2), "]"
    stop 1
end if
if ((c(3)) /= 2) then
    print *, "FAIL: want [2] got [", c(3), "]"
    stop 1
end if
end program t
