! vybe-test: fortran/array_transforms/spread_dim1_length4_copies_3
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(4)=[1,2,3,4]
integer :: m(3,4)
m=spread(a,1,3)
if ((m(2,3)) /= 3) then
    print *, "FAIL: want [3] got [", m(2,3), "]"
    stop 1
end if
if ((m(3,4)) /= 4) then
    print *, "FAIL: want [4] got [", m(3,4), "]"
    stop 1
end if
if ((sum(m)) /= 30) then
    print *, "FAIL: want [30] got [", sum(m), "]"
    stop 1
end if
end program t
