! vybe-test: fortran/character_scan_verify_extended/adjustl_then_len_trim
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=6) :: s = '   z  '
if ((len_trim(adjustl(s))) /= 1) then
    print *, "FAIL: want [1] got [", len_trim(adjustl(s)), "]"
    stop 1
end if
end program t
