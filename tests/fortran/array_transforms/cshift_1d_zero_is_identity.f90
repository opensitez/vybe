! vybe-test: fortran/array_transforms/cshift_1d_zero_is_identity
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(4)=[9,8,7,6]
integer :: b(4)
b=cshift(a,0)
if ((b(1)) /= 9) then
    print *, "FAIL: want [9] got [", b(1), "]"
    stop 1
end if
if ((b(4)) /= 6) then
    print *, "FAIL: want [6] got [", b(4), "]"
    stop 1
end if
end program t
