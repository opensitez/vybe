! vybe-test: fortran/character_intrinsics_extended/adjustl_left_padded_len_trim_is_five
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=12) :: s = '     Fortran'
if ((len_trim(adjustl(s))) /= 7) then
    print *, "FAIL: want [7] got [", len_trim(adjustl(s)), "]"
    stop 1
end if
end program t
