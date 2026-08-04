! vybe-test: fortran/character_intrinsics_extended/len_trim_all_blanks_is_zero
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=6) :: s = '      '
if ((len_trim(s)) /= 0) then
    print *, "FAIL: want [0] got [", len_trim(s), "]"
    stop 1
end if
if ((len(s)) /= 6) then
    print *, "FAIL: want [6] got [", len(s), "]"
    stop 1
end if
end program t
