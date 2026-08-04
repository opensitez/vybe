! vybe-test: fortran/array_reduction_extended/sum_dim1_two_by_three_cols
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])
integer :: c(3)
c = sum(m, dim=1)
if ((c(1)) /= 5) then
    print *, "FAIL: want [5] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= 7) then
    print *, "FAIL: want [7] got [", c(2), "]"
    stop 1
end if
if ((c(3)) /= 9) then
    print *, "FAIL: want [9] got [", c(3), "]"
    stop 1
end if
end program t
