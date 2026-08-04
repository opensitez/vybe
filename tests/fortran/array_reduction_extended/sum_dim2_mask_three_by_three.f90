! vybe-test: fortran/array_reduction_extended/sum_dim2_mask_three_by_three
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(3,3) = reshape([(i, i = 1, 9)],[3,3])
logical :: mask(3,3)
mask = m > 4
integer :: r(3)
r = sum(m, dim=2, mask=mask)
if ((r(1)) /= 0) then
    print *, "FAIL: want [0] got [", r(1), "]"
    stop 1
end if
if ((r(3)) /= 30) then
    print *, "FAIL: want [30] got [", r(3), "]"
    stop 1
end if
end program t
