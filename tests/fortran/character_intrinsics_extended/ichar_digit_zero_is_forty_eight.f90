! vybe-test: fortran/character_intrinsics_extended/ichar_digit_zero_is_forty_eight
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
if ((ichar('0')) /= 48) then
    print *, "FAIL: want [48] got [", ichar('0'), "]"
    stop 1
end if
end program t
