! vybe-test: fortran/character_intrinsics_extended/verify_pure_alpha_returns_zero
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=6) :: s = 'Fortran'
if ((verify(s, 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ')) /= 0) then
    print *, "FAIL: want [0] got [", verify(s, 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'), "]"
    stop 1
end if
end program t
