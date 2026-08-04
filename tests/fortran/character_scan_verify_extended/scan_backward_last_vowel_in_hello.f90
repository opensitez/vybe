! vybe-test: fortran/character_scan_verify_extended/scan_backward_last_vowel_in_hello
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=5) :: s = 'hello'
if ((scan(s, 'aeiou', .true.)) /= 5) then
    print *, "FAIL: want [5] got [", scan(s, 'aeiou', .true.), "]"
    stop 1
end if
end program t
