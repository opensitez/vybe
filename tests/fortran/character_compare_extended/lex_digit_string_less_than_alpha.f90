! vybe-test: fortran/character_compare_extended/lex_digit_string_less_than_alpha
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('999', 'aaa')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('999', 'aaa'), "]"
    stop 1
end if
end program t
