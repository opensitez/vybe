! vybe-test: fortran/array_transforms/spread_dim1_five_by_two_sum
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(2)=[2,3]
integer :: m(5,2)
m=spread(a,1,5)
if ((m(5,2)) /= 3) then
    print *, "FAIL: want [3] got [", m(5,2), "]"
    stop 1
end if
if ((sum(m)) /= 25) then
    print *, "FAIL: want [25] got [", sum(m), "]"
    stop 1
end if
end program t
