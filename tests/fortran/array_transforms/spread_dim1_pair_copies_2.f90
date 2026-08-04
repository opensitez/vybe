! vybe-test: fortran/array_transforms/spread_dim1_pair_copies_2
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(2)=[3,5]
integer :: m(2,2)
m=spread(a,1,2)
if ((m(1,1)) /= 3) then
    print *, "FAIL: want [3] got [", m(1,1), "]"
    stop 1
end if
if ((m(2,2)) /= 5) then
    print *, "FAIL: want [5] got [", m(2,2), "]"
    stop 1
end if
if ((sum(m)) /= 16) then
    print *, "FAIL: want [16] got [", sum(m), "]"
    stop 1
end if
end program t
