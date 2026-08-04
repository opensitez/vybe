! vybe-test: fortran/array_reduction_extended/sum_dim2_two_by_three_rows
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])
integer :: r(2)
r = sum(m, dim=2)
if ((r(1)) /= 6) then
    print *, "FAIL: want [6] got [", r(1), "]"
    stop 1
end if
if ((r(2)) /= 15) then
    print *, "FAIL: want [15] got [", r(2), "]"
    stop 1
end if
end program t
