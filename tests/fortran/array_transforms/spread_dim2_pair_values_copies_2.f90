! vybe-test: fortran/array_transforms/spread_dim2_pair_values_copies_2
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(2)=[7,9]
integer :: m(2,2)
m=spread(a,2,2)
if ((m(1,1)) /= 7) then
    print *, "FAIL: want [7] got [", m(1,1), "]"
    stop 1
end if
if ((m(2,2)) /= 9) then
    print *, "FAIL: want [9] got [", m(2,2), "]"
    stop 1
end if
if ((sum(m)) /= 32) then
    print *, "FAIL: want [32] got [", sum(m), "]"
    stop 1
end if
end program t
