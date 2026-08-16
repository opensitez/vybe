! vybe-test: fortran/array_reduction_extended/sum_real_dim1_mask_fractions
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
real :: m(2,3) = reshape([1.0,2.0,3.0,4.0,5.0,6.0],[2,3])
logical :: mask(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])
real :: c(3)
c = sum(m, dim=1, mask=mask)
if ((c(1)) /= 1.00000000) then
    print *, "FAIL: want [1.00000000] got [", c(1), "]"
    stop 1
end if
if ((c(3)) /= 5.00000000) then
    print *, "FAIL: want [5.00000000] got [", c(3), "]"
    stop 1
end if
end program t
