! vybe-test: fortran/array_transforms/spread_dim1_zeros_copies_3
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(1)=[0]
integer :: m(3,1)
m=spread(a,1,3)
if ((m(2,1)) /= 0) then
    print *, "FAIL: want [0] got [", m(2,1), "]"
    stop 1
end if
if ((sum(m)) /= 0) then
    print *, "FAIL: want [0] got [", sum(m), "]"
    stop 1
end if
end program t
