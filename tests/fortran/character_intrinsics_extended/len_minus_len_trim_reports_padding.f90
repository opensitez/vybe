! vybe-test: fortran/character_intrinsics_extended/len_minus_len_trim_reports_padding
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=15) :: s = 'payload'
if ((len(s) - len_trim(s)) /= 8) then
    print *, "FAIL: want [8] got [", len(s) - len_trim(s), "]"
    stop 1
end if
end program t
