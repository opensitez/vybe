! vybe-test: fortran/character_compare_extended/lge_beta_greater_or_equal_alpha
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lge('beta', 'alpha')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lge('beta', 'alpha'), "]"
    stop 1
end if
end program t
