! vybe-test: fortran/character_compare_extended/lex_space_less_than_digit
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt(' ', '0')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt(' ', '0'), "]"
    stop 1
end if
end program t
