! vybe-test: fortran/array_transforms/spread_dim2_zeros_copies_4
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(2)=[0,0]
integer :: m(2,4)
m=spread(a,2,4)
if ((m(1,1)) /= 0) then
    print *, "FAIL: want [0] got [", m(1,1), "]"
    stop 1
end if
if ((sum(m)) /= 0) then
    print *, "FAIL: want [0] got [", sum(m), "]"
    stop 1
end if
end program t
