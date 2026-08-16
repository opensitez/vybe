! vybe-test: fortran/array_transforms/cshift_1d_length4_rotate_three
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(4)=[1,2,3,4]
integer :: b(4)
b=cshift(a,3)
if ((b(1)) /= 4) then
    print *, "FAIL: want [4] got [", b(1), "]"
    stop 1
end if
if ((b(4)) /= 3) then
    print *, "FAIL: want [3] got [", b(4), "]"
    stop 1
end if
end program t
