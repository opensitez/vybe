! vybe-test: fortran/character_scan_verify_extended/verify_all_alnum_set
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=9) :: s = 'Fortran90'
if ((verify(s, 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789')) /= 0) then
    print *, "FAIL: want [0] got [", verify(s, 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'), "]"
    stop 1
end if
end program t
