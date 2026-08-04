! vybe-test: fortran/character_scan_verify_extended/verify_leading_blank_not_in_set
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=6) :: s = '  data'
if ((verify(s, 'abcdefghijklmnopqrstuvwxyz')) /= 1) then
    print *, "FAIL: want [1] got [", verify(s, 'abcdefghijklmnopqrstuvwxyz'), "]"
    stop 1
end if
end program t
