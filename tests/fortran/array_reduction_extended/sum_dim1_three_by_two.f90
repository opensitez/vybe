! vybe-test: fortran/array_reduction_extended/sum_dim1_three_by_two
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(3,2) = reshape([1,2,3,4,5,6],[3,2])
integer :: c(2)
c = sum(m, dim=1)
if ((c(1)) /= 9) then
    print *, "FAIL: want [9] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= 12) then
    print *, "FAIL: want [12] got [", c(2), "]"
    stop 1
end if
end program t
