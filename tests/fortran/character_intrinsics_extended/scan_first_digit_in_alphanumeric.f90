! vybe-test: fortran/character_intrinsics_extended/scan_first_digit_in_alphanumeric
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=6) :: s = 'abc123'
if ((scan(s, '0123456789')) /= 4) then
    print *, "FAIL: want [4] got [", scan(s, '0123456789'), "]"
    stop 1
end if
end program t
