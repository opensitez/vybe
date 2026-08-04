! vybe-test: fortran/character_intrinsics_extended/index_back_finds_last_xy_pair
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=7) :: s = 'xyzzyxy'
if ((index(s, 'xy', .true.)) /= 6) then
    print *, "FAIL: want [6] got [", index(s, 'xy', .true.), "]"
    stop 1
end if
end program t
