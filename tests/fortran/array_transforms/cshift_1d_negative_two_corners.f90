! vybe-test: fortran/array_transforms/cshift_1d_negative_two_corners
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(6)=[1,2,3,4,5,6]
integer :: b(6)
b=cshift(a,-2)
if ((b(1)) /= 5) then
    print *, "FAIL: want [5] got [", b(1), "]"
    stop 1
end if
if ((b(4)) /= 2) then
    print *, "FAIL: want [2] got [", b(4), "]"
    stop 1
end if
end program t
