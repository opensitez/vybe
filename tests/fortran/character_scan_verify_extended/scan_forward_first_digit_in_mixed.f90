! vybe-test: fortran/character_scan_verify_extended/scan_forward_first_digit_in_mixed
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=9) :: s = 'abc123def'
if ((scan(s, '0123456789')) /= 4) then
    print *, "FAIL: want [4] got [", scan(s, '0123456789'), "]"
    stop 1
end if
end program t
