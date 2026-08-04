! vybe-test: fortran/array_transforms/spread_dim1_three_rows_sum
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3)=[4,5,6]
integer :: m(3,3)
m=spread(a,1,3)
if ((m(3,2)) /= 5) then
    print *, "FAIL: want [5] got [", m(3,2), "]"
    stop 1
end if
if ((sum(m)) /= 45) then
    print *, "FAIL: want [45] got [", sum(m), "]"
    stop 1
end if
end program t
