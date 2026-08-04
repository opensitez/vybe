! vybe-test: fortran/character_intrinsics_extended/index_back_single_letter_last_position
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=8) :: s = 'abracada'
if ((index(s, 'a', .true.)) /= 8) then
    print *, "FAIL: want [8] got [", index(s, 'a', .true.), "]"
    stop 1
end if
end program t
