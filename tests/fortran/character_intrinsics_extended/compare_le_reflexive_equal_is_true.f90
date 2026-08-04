! vybe-test: fortran/character_intrinsics_extended/compare_le_reflexive_equal_is_true
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
if (('pair' <= 'pair') .neqv. .true.) then
    print *, "FAIL: want [true] got [", 'pair' <= 'pair', "]"
    stop 1
end if
end program t
