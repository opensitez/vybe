! vybe-test: fortran/character_scan_verify_extended/index_backward_suffix
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=13) :: s = 'prefix-suffix'
if ((index(s, 'suffix')) /= 8) then
    print *, "FAIL: want [8] got [", index(s, 'suffix'), "]"
    stop 1
end if
end program t
