! vybe-test: fortran/array_transforms/spread_dim2_singleton_five_copies_3
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(1)=[5]
integer :: m(1,3)
m=spread(a,2,3)
if ((m(1,1)) /= 5) then
    print *, "FAIL: want [5] got [", m(1,1), "]"
    stop 1
end if
if ((m(1,3)) /= 5) then
    print *, "FAIL: want [5] got [", m(1,3), "]"
    stop 1
end if
if ((sum(m)) /= 15) then
    print *, "FAIL: want [15] got [", sum(m), "]"
    stop 1
end if
end program t
