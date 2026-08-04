! vybe-test: fortran/character_intrinsics_extended/scan_no_vowel_returns_zero
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=5) :: s = 'rhythm'
if ((scan(s, 'aeiou')) /= 0) then
    print *, "FAIL: want [0] got [", scan(s, 'aeiou'), "]"
    stop 1
end if
end program t
