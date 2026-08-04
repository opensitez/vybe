! vybe-test: fortran/character_intrinsics_extended/verify_all_alnum_chars_in_set
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=6) :: s = 'A1b2C3'
if ((verify(s, '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz')) /= 0) then
    print *, "FAIL: want [0] got [", verify(s, '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'), "]"
    stop 1
end if
end program t
