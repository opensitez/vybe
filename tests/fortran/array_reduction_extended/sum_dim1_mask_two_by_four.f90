! vybe-test: fortran/array_reduction_extended/sum_dim1_mask_two_by_four
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(2,4) = reshape([1,2,3,4,5,6,7,8],[2,4])
logical :: mask(2,4) = reshape([.true.,.false.,.true.,.false.,.true.,.false.,.true.,.false.],[2,4])
integer :: c(4)
c = sum(m, dim=1, mask=mask)
if ((c(1)) /= 1) then
    print *, "FAIL: want [1] got [", c(1), "]"
    stop 1
end if
if ((c(3)) /= 5) then
    print *, "FAIL: want [5] got [", c(3), "]"
    stop 1
end if
end program t
