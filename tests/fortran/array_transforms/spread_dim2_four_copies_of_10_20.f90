! vybe-test: fortran/array_transforms/spread_dim2_four_copies_of_10_20
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(2)=[10,20]
integer :: m(2,4)
m=spread(a,2,4)
if ((m(1,4)) /= 10) then
    print *, "FAIL: want [10] got [", m(1,4), "]"
    stop 1
end if
if ((m(2,3)) /= 20) then
    print *, "FAIL: want [20] got [", m(2,3), "]"
    stop 1
end if
if ((sum(m)) /= 120) then
    print *, "FAIL: want [120] got [", sum(m), "]"
    stop 1
end if
end program t
