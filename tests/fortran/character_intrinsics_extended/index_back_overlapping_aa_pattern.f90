! vybe-test: fortran/character_intrinsics_extended/index_back_overlapping_aa_pattern
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=4) :: s = 'baaa'
if ((index(s, 'aa', .true.)) /= 3) then
    print *, "FAIL: want [3] got [", index(s, 'aa', .true.), "]"
    stop 1
end if
end program t
