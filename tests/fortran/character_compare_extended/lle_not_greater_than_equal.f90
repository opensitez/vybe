! vybe-test: fortran/character_compare_extended/lle_not_greater_than_equal
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lle('beta', 'alpha')) .neqv. .false.) then
    print *, "FAIL: want [false] got [", lle('beta', 'alpha'), "]"
    stop 1
end if
end program t
