! vybe-test: fortran/character_intrinsics_extended/scan_back_last_vowel_in_word
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=9) :: s = 'rhythmics'
if ((scan(s, 'aeiou', .true.)) /= 7) then
    print *, "FAIL: want [7] got [", scan(s, 'aeiou', .true.), "]"
    stop 1
end if
end program t
