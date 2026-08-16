! vybe-test: fortran/array_reduction_extended/sum_dim1_three_by_two
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(3,2) = reshape([1,2,3,4,5,6],[3,2])
integer :: c(2)
c = sum(m, dim=1)
if ((c(1)) /= 6) then
    print *, "FAIL: want [6] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= 15) then
    print *, "FAIL: want [15] got [", c(2), "]"
    stop 1
end if
end program t
