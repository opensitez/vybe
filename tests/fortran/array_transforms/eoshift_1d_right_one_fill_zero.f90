! vybe-test: fortran/array_transforms/eoshift_1d_right_one_fill_zero
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(4)=[1,2,3,4]
integer :: b(4)
b=eoshift(a,-1)
if ((b(1)) /= 0) then
    print *, "FAIL: want [0] got [", b(1), "]"
    stop 1
end if
if ((b(4)) /= 3) then
    print *, "FAIL: want [3] got [", b(4), "]"
    stop 1
end if
end program t
