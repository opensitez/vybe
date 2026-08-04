! vybe-test: fortran/array_transforms/merge_scalar_true_picks_first
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: x
x=merge(42,99,.true.)
if ((x) /= 42) then
    print *, "FAIL: want [42] got [", x, "]"
    stop 1
end if
end program t
