! vybe-test: fortran/array_transforms/spread_dim2_length4_copies_3_sum
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(4)=[1,2,3,4]
integer :: m(4,3)
m=spread(a,2,3)
if ((m(3,1)) /= 3) then
    print *, "FAIL: want [3] got [", m(3,1), "]"
    stop 1
end if
if ((m(4,3)) /= 4) then
    print *, "FAIL: want [4] got [", m(4,3), "]"
    stop 1
end if
if ((sum(m)) /= 30) then
    print *, "FAIL: want [30] got [", sum(m), "]"
    stop 1
end if
end program t
