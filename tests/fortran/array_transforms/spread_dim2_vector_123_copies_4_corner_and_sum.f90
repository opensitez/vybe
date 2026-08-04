! vybe-test: fortran/array_transforms/spread_dim2_vector_123_copies_4_corner_and_sum
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3)=[1,2,3]
integer :: m(3,4)
m=spread(a,2,4)
if ((m(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", m(1,1), "]"
    stop 1
end if
if ((m(2,4)) /= 2) then
    print *, "FAIL: want [2] got [", m(2,4), "]"
    stop 1
end if
if ((sum(m)) /= 24) then
    print *, "FAIL: want [24] got [", sum(m), "]"
    stop 1
end if
end program t
