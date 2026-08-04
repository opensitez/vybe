! vybe-test: fortran/character_scan_verify_extended/index_backward_last_bc
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=6) :: s = 'abcabc'
if ((index(s, 'bc')) /= 5) then
    print *, "FAIL: want [5] got [", index(s, 'bc'), "]"
    stop 1
end if
end program t
