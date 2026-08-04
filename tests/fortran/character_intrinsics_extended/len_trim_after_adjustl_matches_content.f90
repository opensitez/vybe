! vybe-test: fortran/character_intrinsics_extended/len_trim_after_adjustl_matches_content
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=12) :: s = '   data'
if ((len_trim(adjustl(s))) /= 4) then
    print *, "FAIL: want [4] got [", len_trim(adjustl(s)), "]"
    stop 1
end if
if ((len(s)) /= 12) then
    print *, "FAIL: want [12] got [", len(s), "]"
    stop 1
end if
end program t
