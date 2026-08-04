! vybe-test: fortran/array_transforms/cshift_1d_left_two_first_second
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(5)=[1,2,3,4,5]
integer :: b(5)
b=cshift(a,2)
if ((b(1)) /= 3) then
    print *, "FAIL: want [3] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 4) then
    print *, "FAIL: want [4] got [", b(2), "]"
    stop 1
end if
end program t
