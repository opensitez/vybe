! vybe-test: fortran/character_scan_verify_extended/verify_space_is_nonalpha
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=5) :: s = 'ab cd'
if ((verify(s, 'abcdefghijklmnopqrstuvwxyz')) /= 3) then
    print *, "FAIL: want [3] got [", verify(s, 'abcdefghijklmnopqrstuvwxyz'), "]"
    stop 1
end if
end program t
