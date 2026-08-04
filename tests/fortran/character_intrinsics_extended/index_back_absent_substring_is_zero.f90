! vybe-test: fortran/character_intrinsics_extended/index_back_absent_substring_is_zero
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=10) :: s = 'fortran 90'
if ((index(s, 'cpp', .true.)) /= 0) then
    print *, "FAIL: want [0] got [", index(s, 'cpp', .true.), "]"
    stop 1
end if
end program t
