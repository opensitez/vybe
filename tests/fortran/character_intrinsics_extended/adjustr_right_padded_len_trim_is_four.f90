! vybe-test: fortran/character_intrinsics_extended/adjustr_right_padded_len_trim_is_four
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=12) :: s = 'Code'
if ((len_trim(adjustr(s))) /= 12) then
    print *, "FAIL: want [12] got [", len_trim(adjustr(s)), "]"
    stop 1
end if
end program t
