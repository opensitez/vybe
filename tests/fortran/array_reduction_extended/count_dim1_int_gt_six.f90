! vybe-test: fortran/array_reduction_extended/count_dim1_int_gt_six
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(3,4) = reshape([(i, i = 1, 12)],[3,4])
integer :: c(4)
c = count(m > 6, dim=1)
if ((c(1)) /= 1) then
    print *, "FAIL: want [1] got [", c(1), "]"
    stop 1
end if
if ((c(3)) /= 2) then
    print *, "FAIL: want [2] got [", c(3), "]"
    stop 1
end if
if ((c(4)) /= 2) then
    print *, "FAIL: want [2] got [", c(4), "]"
    stop 1
end if
end program t
