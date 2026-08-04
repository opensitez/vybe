! vybe-test: fortran/array_transforms/merge_scalar_negative_branch
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: x
x=merge(-3,7,.false.)
if ((x) /= 7) then
    print *, "FAIL: want [7] got [", x, "]"
    stop 1
end if
end program t
