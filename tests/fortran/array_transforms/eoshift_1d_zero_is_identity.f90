! vybe-test: fortran/array_transforms/eoshift_1d_zero_is_identity
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(4)=[4,5,6,7]
integer :: b(4)
b=eoshift(a,0)
if ((b(1)) /= 4) then
    print *, "FAIL: want [4] got [", b(1), "]"
    stop 1
end if
if ((b(4)) /= 7) then
    print *, "FAIL: want [7] got [", b(4), "]"
    stop 1
end if
end program t
