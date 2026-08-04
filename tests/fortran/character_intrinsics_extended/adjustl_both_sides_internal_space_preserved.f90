! vybe-test: fortran/character_intrinsics_extended/adjustl_both_sides_internal_space_preserved
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=16) :: s = '   ab cd   '
if (trim(trim(adjustl(s))) /= "ab cd") then
    print *, "FAIL: want [ab cd] got [", trim(adjustl(s)), "]"
    stop 1
end if
end program t
