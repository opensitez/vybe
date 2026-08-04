! vybe-test: fortran/character_compare_extended/lex_case_pair_lower_greater_upper
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lgt('dog', 'DOG')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lgt('dog', 'DOG'), "]"
    stop 1
end if
end program t
