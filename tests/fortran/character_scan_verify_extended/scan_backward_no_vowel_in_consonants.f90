! vybe-test: fortran/character_scan_verify_extended/scan_backward_no_vowel_in_consonants
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=5) :: s = 'bcdfg'
if ((scan(s, 'aeiou', .true.)) /= 0) then
    print *, "FAIL: want [0] got [", scan(s, 'aeiou', .true.), "]"
    stop 1
end if
end program t
