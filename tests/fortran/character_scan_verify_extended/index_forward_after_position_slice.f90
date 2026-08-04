! vybe-test: fortran/character_scan_verify_extended/index_forward_after_position_slice
! origin: languages/fortran/tests/fortran/test_character_scan_verify_extended.rs
program t
character(len=11) :: s = 'one two one'
if ((index(s(5:), 'one')) /= 5) then
    print *, "FAIL: want [5] got [", index(s(5:), 'one'), "]"
    stop 1
end if
end program t
