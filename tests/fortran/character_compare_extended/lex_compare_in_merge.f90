! vybe-test: fortran/character_compare_extended/lex_compare_in_merge
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((merge(1, 0, llt('a', 'b'))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, llt('a', 'b')), "]"
    stop 1
end if
end program t
