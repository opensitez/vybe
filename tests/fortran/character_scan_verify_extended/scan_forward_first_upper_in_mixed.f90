! vybe-test: fortran/character_scan_verify_extended/scan_forward_first_upper_in_mixed
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=6) :: s = 'aBcDeF'
if ((scan(s, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ')) /= 2) then
    print *, "FAIL: want [2] got [", scan(s, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), "]"
    stop 1
end if
end program t
