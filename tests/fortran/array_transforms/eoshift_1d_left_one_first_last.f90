! vybe-test: fortran/array_transforms/eoshift_1d_left_one_first_last
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(5)=[1,2,3,4,5]
integer :: b(5)
b=eoshift(a,1)
if ((b(1)) /= 2) then
    print *, "FAIL: want [2] got [", b(1), "]"
    stop 1
end if
if ((b(4)) /= 5) then
    print *, "FAIL: want [5] got [", b(4), "]"
    stop 1
end if
end program t
