! vybe-test: fortran/array_transforms/eoshift_1d_left_one_length_three
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3)=[10,20,30]
integer :: b(3)
b=eoshift(a,1)
if ((b(1)) /= 20) then
    print *, "FAIL: want [20] got [", b(1), "]"
    stop 1
end if
if ((b(3)) /= 30) then
    print *, "FAIL: want [30] got [", b(3), "]"
    stop 1
end if
end program t
