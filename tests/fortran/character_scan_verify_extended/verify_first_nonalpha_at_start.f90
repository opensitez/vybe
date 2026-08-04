! vybe-test: fortran/character_scan_verify_extended/verify_first_nonalpha_at_start
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=4) :: s = '1abc'
if ((verify(s, '0123456789')) /= 1) then
    print *, "FAIL: want [1] got [", verify(s, '0123456789'), "]"
    stop 1
end if
end program t
