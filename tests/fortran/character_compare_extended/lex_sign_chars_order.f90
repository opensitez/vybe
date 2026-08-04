! vybe-test: fortran/character_compare_extended/lex_sign_chars_order
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('+', '-')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('+', '-'), "]"
    stop 1
end if
end program t
