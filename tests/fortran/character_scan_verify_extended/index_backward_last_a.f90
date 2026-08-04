! vybe-test: fortran/character_scan_verify_extended/index_backward_last_a
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=8) :: s = 'abracada'
if ((index(s, 'a')) /= 8) then
    print *, "FAIL: want [8] got [", index(s, 'a'), "]"
    stop 1
end if
end program t
