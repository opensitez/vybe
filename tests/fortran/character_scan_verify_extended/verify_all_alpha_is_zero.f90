! vybe-test: fortran/character_scan_verify_extended/verify_all_alpha_is_zero
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=8) :: s = 'alphabet'
if ((verify(s, 'abcdefghijklmnopqrstuvwxyz')) /= 0) then
    print *, "FAIL: want [0] got [", verify(s, 'abcdefghijklmnopqrstuvwxyz'), "]"
    stop 1
end if
end program t
