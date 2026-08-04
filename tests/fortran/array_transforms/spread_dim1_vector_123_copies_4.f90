! vybe-test: fortran/array_transforms/spread_dim1_vector_123_copies_4
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3)=[1,2,3]
integer :: m(4,3)
m=spread(a,1,4)
if ((m(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", m(1,1), "]"
    stop 1
end if
if ((m(4,2)) /= 2) then
    print *, "FAIL: want [2] got [", m(4,2), "]"
    stop 1
end if
if ((sum(m)) /= 24) then
    print *, "FAIL: want [24] got [", sum(m), "]"
    stop 1
end if
end program t
