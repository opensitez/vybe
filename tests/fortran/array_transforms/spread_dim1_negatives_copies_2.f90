! vybe-test: fortran/array_transforms/spread_dim1_negatives_copies_2
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3)=[-2,1,4]
integer :: m(2,3)
m=spread(a,1,2)
if ((m(1,2)) /= 1) then
    print *, "FAIL: want [1] got [", m(1,2), "]"
    stop 1
end if
if ((m(2,3)) /= 4) then
    print *, "FAIL: want [4] got [", m(2,3), "]"
    stop 1
end if
if ((sum(m)) /= 6) then
    print *, "FAIL: want [6] got [", sum(m), "]"
    stop 1
end if
end program t
