! vybe-test: fortran/array_reduction_extended/count_dim1_mask_comparison
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(2,4) = reshape([1,5,2,8,3,6,4,7],[2,4])
logical :: mask(2,4) = reshape([.true.,.true.,.false.,.true.,.false.,.true.,.true.,.false.],[2,4])
integer :: c(4)
c = count(m > 4, dim=1, mask=mask)
if ((c(1)) /= 1) then
    print *, "FAIL: want [1] got [", c(1), "]"
    stop 1
end if
if ((c(4)) /= 1) then
    print *, "FAIL: want [1] got [", c(4), "]"
    stop 1
end if
end program t
