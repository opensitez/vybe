! vybe-test: fortran/array_transforms/spread_dim1_singleton_copies_3
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(1)=[8]
integer :: m(3,1)
m=spread(a,1,3)
if ((m(1,1)) /= 8) then
    print *, "FAIL: want [8] got [", m(1,1), "]"
    stop 1
end if
if ((m(3,1)) /= 8) then
    print *, "FAIL: want [8] got [", m(3,1), "]"
    stop 1
end if
if ((sum(m)) /= 24) then
    print *, "FAIL: want [24] got [", sum(m), "]"
    stop 1
end if
end program t
