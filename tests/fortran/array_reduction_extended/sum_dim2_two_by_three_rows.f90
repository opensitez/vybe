! vybe-test: fortran/array_reduction_extended/sum_dim2_two_by_three_rows
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])
integer :: r(2)
r = sum(m, dim=2)
if ((r(1)) /= 9) then
    print *, "FAIL: want [9] got [", r(1), "]"
    stop 1
end if
if ((r(2)) /= 12) then
    print *, "FAIL: want [12] got [", r(2), "]"
    stop 1
end if
end program t
