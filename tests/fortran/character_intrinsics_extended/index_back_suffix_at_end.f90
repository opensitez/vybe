! vybe-test: fortran/character_intrinsics_extended/index_back_suffix_at_end
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=12) :: s = 'prefix-suffix'
if ((index(s, 'suffix', .true.)) /= 8) then
    print *, "FAIL: want [8] got [", index(s, 'suffix', .true.), "]"
    stop 1
end if
end program t
