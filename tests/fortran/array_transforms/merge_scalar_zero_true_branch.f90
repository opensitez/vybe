! vybe-test: fortran/array_transforms/merge_scalar_zero_true_branch
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: x
x=merge(0,5,.true.)
if ((x) /= 0) then
    print *, "FAIL: want [0] got [", x, "]"
    stop 1
end if
end program t
