! vybe-test: fortran/array_transforms/eoshift_1d_left_two_boundary
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(5)=[1,2,3,4,5]
integer :: b(5)
b=eoshift(a,2,-9)
if ((b(3)) /= 5) then
    print *, "FAIL: want [5] got [", b(3), "]"
    stop 1
end if
if ((b(4)) /= -9) then
    print *, "FAIL: want [-9] got [", b(4), "]"
    stop 1
end if
end program t
