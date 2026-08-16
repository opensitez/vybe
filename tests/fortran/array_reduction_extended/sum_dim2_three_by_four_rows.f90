! vybe-test: fortran/array_reduction_extended/sum_dim2_three_by_four_rows
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(3,4) = reshape([(i, i = 1, 12)],[3,4])
integer :: r(3)
r = sum(m, dim=2)
if ((r(1)) /= 22) then
    print *, "FAIL: want [22] got [", r(1), "]"
    stop 1
end if
if ((r(2)) /= 26) then
    print *, "FAIL: want [26] got [", r(2), "]"
    stop 1
end if
if ((r(3)) /= 30) then
    print *, "FAIL: want [30] got [", r(3), "]"
    stop 1
end if
end program t
