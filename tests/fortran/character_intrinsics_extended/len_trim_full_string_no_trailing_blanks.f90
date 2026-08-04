! vybe-test: fortran/character_intrinsics_extended/len_trim_full_string_no_trailing_blanks
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=5) :: s = 'abcde'
if ((len(s)) /= 5) then
    print *, "FAIL: want [5] got [", len(s), "]"
    stop 1
end if
if ((len_trim(s)) /= 5) then
    print *, "FAIL: want [5] got [", len_trim(s), "]"
    stop 1
end if
end program t
