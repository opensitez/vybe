! vybe-test: fortran/character_scan_verify_extended/scan_backward_last_space_in_padded
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=7) :: s = 'a b c  '
if ((scan(s, ' ', .true.)) /= 5) then
    print *, "FAIL: want [5] got [", scan(s, ' ', .true.), "]"
    stop 1
end if
end program t
