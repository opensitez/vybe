! vybe-test: fortran/character_intrinsics_extended/verify_symbol_outside_alnum_set
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=6) :: s = 'ok!now'
if ((verify(s, 'abcdefghijklmnopqrstuvwxyz')) /= 3) then
    print *, "FAIL: want [3] got [", verify(s, 'abcdefghijklmnopqrstuvwxyz'), "]"
    stop 1
end if
end program t
