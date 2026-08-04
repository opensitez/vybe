! vybe-test: fortran/character_intrinsics_extended/iachar_space_is_thirty_two
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
if ((iachar(' ')) /= 32) then
    print *, "FAIL: want [32] got [", iachar(' '), "]"
    stop 1
end if
end program t
