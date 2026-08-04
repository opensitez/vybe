! vybe-test: fortran/character_scan_verify_extended/scan_backward_last_digit_in_trailing
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=7) :: s = 'test007'
if ((scan(s, '0123456789', .true.)) /= 7) then
    print *, "FAIL: want [7] got [", scan(s, '0123456789', .true.), "]"
    stop 1
end if
end program t
