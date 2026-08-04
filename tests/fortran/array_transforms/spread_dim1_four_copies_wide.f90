! vybe-test: fortran/array_transforms/spread_dim1_four_copies_wide
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(2)=[11,13]
integer :: m(4,2)
m=spread(a,1,4)
if ((m(4,1)) /= 11) then
    print *, "FAIL: want [11] got [", m(4,1), "]"
    stop 1
end if
if ((m(2,2)) /= 13) then
    print *, "FAIL: want [13] got [", m(2,2), "]"
    stop 1
end if
if ((sum(m)) /= 96) then
    print *, "FAIL: want [96] got [", sum(m), "]"
    stop 1
end if
end program t
