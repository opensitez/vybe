! vybe-test: fortran/character_intrinsics_extended/verify_space_not_in_alpha_set
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=7) :: s = 'ab cd ef'
if ((verify(s, 'abcdefghijklmnopqrstuvwxyz')) /= 3) then
    print *, "FAIL: want [3] got [", verify(s, 'abcdefghijklmnopqrstuvwxyz'), "]"
    stop 1
end if
end program t
