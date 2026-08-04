! vybe-test: fortran/character_intrinsics_extended/adjustl_all_blanks_becomes_blank
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=8) :: s = '        '
if ((len_trim(adjustl(s))) /= 0) then
    print *, "FAIL: want [0] got [", len_trim(adjustl(s)), "]"
    stop 1
end if
end program t
