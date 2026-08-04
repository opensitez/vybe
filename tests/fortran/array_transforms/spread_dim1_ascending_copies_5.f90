! vybe-test: fortran/array_transforms/spread_dim1_ascending_copies_5
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(2)=[1,9]
integer :: m(5,2)
m=spread(a,1,5)
if ((m(5,1)) /= 1) then
    print *, "FAIL: want [1] got [", m(5,1), "]"
    stop 1
end if
if ((m(1,2)) /= 9) then
    print *, "FAIL: want [9] got [", m(1,2), "]"
    stop 1
end if
if ((sum(m)) /= 50) then
    print *, "FAIL: want [50] got [", sum(m), "]"
    stop 1
end if
end program t
