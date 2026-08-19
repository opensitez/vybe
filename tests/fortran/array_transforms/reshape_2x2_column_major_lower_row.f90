! vybe-test: fortran/array_transforms/reshape_2x2_column_major_lower_row
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(4)=[1,2,3,4]
integer :: m(2,2)
m=reshape(a,[2,2])
if ((m(2,1)) /= 2) then
    print *, "FAIL: want [2] got [", m(2,1), "]"
    stop 1
end if
if ((m(2,2)) /= 4) then
    print *, "FAIL: want [4] got [", m(2,2), "]"
    stop 1
end if
end program t
