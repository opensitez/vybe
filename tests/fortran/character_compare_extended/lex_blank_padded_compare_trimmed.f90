! vybe-test: fortran/character_compare_extended/lex_blank_padded_compare_trimmed
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
character(len=5) :: a = 'hi   '
character(len=5) :: b = 'hi'
if ((lge(a, b)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lge(a, b), "]"
    stop 1
end if
end program t
