! vybe-test: fortran/array_transforms/spread_dim2_five_elements_two_copies
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(5)=[1,1,2,2,3]
integer :: m(5,2)
m=spread(a,2,2)
if ((m(5,1)) /= 3) then
    print *, "FAIL: want [3] got [", m(5,1), "]"
    stop 1
end if
if ((m(1,2)) /= 1) then
    print *, "FAIL: want [1] got [", m(1,2), "]"
    stop 1
end if
if ((sum(m)) /= 18) then
    print *, "FAIL: want [18] got [", sum(m), "]"
    stop 1
end if
end program t
