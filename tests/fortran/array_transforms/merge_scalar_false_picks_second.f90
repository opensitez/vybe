! vybe-test: fortran/array_transforms/merge_scalar_false_picks_second
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: x
x=merge(42,99,.false.)
if ((x) /= 99) then
    print *, "FAIL: want [99] got [", x, "]"
    stop 1
end if
end program t
