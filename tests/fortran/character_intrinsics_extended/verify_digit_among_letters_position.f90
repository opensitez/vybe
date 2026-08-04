! vybe-test: fortran/character_intrinsics_extended/verify_digit_among_letters_position
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
character(len=5) :: s = 'ab2de'
if ((verify(s, 'abcde')) /= 3) then
    print *, "FAIL: want [3] got [", verify(s, 'abcde'), "]"
    stop 1
end if
end program t
