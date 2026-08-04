! vybe-test: fortran/array_reduction_extended/product_dim2_three_by_two
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: m(3,2) = reshape([2,3,4,5,6,7],[3,2])
integer :: r(3)
r = product(m, dim=2)
if ((r(1)) /= 6) then
    print *, "FAIL: want [6] got [", r(1), "]"
    stop 1
end if
if ((r(2)) /= 20) then
    print *, "FAIL: want [20] got [", r(2), "]"
    stop 1
end if
if ((r(3)) /= 42) then
    print *, "FAIL: want [42] got [", r(3), "]"
    stop 1
end if
end program t
