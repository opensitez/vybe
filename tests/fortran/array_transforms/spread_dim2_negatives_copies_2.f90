! vybe-test: fortran/array_transforms/spread_dim2_negatives_copies_2
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3)=[-1,0,2]
integer :: m(3,2)
m=spread(a,2,2)
if ((m(1,2)) /= -1) then
    print *, "FAIL: want [-1] got [", m(1,2), "]"
    stop 1
end if
if ((m(3,1)) /= 2) then
    print *, "FAIL: want [2] got [", m(3,1), "]"
    stop 1
end if
if ((sum(m)) /= 2) then
    print *, "FAIL: want [2] got [", sum(m), "]"
    stop 1
end if
end program t
