! vybe-test: fortran/array_reduction_extended/product_dim2_mask_two_by_three
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(2,3) = reshape([2,3,4,5,6,7],[2,3])
logical :: mask(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])
integer :: r(2)
r = product(m, dim=2, mask=mask)
if ((r(1)) /= 8) then
    print *, "FAIL: want [8] got [", r(1), "]"
    stop 1
end if
if ((r(2)) /= 30) then
    print *, "FAIL: want [30] got [", r(2), "]"
    stop 1
end if
end program t
