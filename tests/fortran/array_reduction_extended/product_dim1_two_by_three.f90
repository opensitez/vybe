! vybe-test: fortran/array_reduction_extended/product_dim1_two_by_three
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])
integer :: c(3)
c = product(m, dim=1)
if ((c(1)) /= 4) then
    print *, "FAIL: want [4] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= 10) then
    print *, "FAIL: want [10] got [", c(2), "]"
    stop 1
end if
if ((c(3)) /= 18) then
    print *, "FAIL: want [18] got [", c(3), "]"
    stop 1
end if
end program t
