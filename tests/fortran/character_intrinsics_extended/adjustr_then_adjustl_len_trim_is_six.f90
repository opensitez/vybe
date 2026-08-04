! vybe-test: fortran/character_intrinsics_extended/adjustr_then_adjustl_len_trim_is_six
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=14) :: s = '  nested  '
if ((len_trim(adjustl(adjustr(s)))) /= 6) then
    print *, "FAIL: want [6] got [", len_trim(adjustl(adjustr(s))), "]"
    stop 1
end if
end program t
