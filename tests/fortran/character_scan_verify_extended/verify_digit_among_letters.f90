! vybe-test: fortran/character_scan_verify_extended/verify_digit_among_letters
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=5) :: s = 'ab2de'
if ((verify(s, 'abcde')) /= 3) then
    print *, "FAIL: want [3] got [", verify(s, 'abcde'), "]"
    stop 1
end if
end program t
