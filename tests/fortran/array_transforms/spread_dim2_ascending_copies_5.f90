! vybe-test: fortran/array_transforms/spread_dim2_ascending_copies_5
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(2)=[4,6]
integer :: m(2,5)
m=spread(a,2,5)
if ((m(1,5)) /= 4) then
    print *, "FAIL: want [4] got [", m(1,5), "]"
    stop 1
end if
if ((m(2,1)) /= 6) then
    print *, "FAIL: want [6] got [", m(2,1), "]"
    stop 1
end if
if ((sum(m)) /= 50) then
    print *, "FAIL: want [50] got [", sum(m), "]"
    stop 1
end if
end program t
