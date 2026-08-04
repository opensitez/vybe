! vybe-test: fortran/array_transforms/eoshift_1d_right_two_boundary
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(5)=[1,2,3,4,5]
integer :: b(5)
b=eoshift(a,-2,0)
if ((b(1)) /= 0) then
    print *, "FAIL: want [0] got [", b(1), "]"
    stop 1
end if
if ((b(3)) /= 1) then
    print *, "FAIL: want [1] got [", b(3), "]"
    stop 1
end if
end program t
