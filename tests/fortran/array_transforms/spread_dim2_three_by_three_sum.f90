! vybe-test: fortran/array_transforms/spread_dim2_three_by_three_sum
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3)=[2,2,2]
integer :: m(3,3)
m=spread(a,2,3)
if ((m(2,2)) /= 2) then
    print *, "FAIL: want [2] got [", m(2,2), "]"
    stop 1
end if
if ((sum(m)) /= 18) then
    print *, "FAIL: want [18] got [", sum(m), "]"
    stop 1
end if
end program t
