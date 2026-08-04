! vybe-test: fortran/character_intrinsics_extended/compare_gt_y_greater_than_x
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
if (('y' > 'x') .neqv. .true.) then
    print *, "FAIL: want [true] got [", 'y' > 'x', "]"
    stop 1
end if
end program t
