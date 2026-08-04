! vybe-test: fortran/character_compare_extended/lgt_beta_after_alpha
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lgt('beta', 'alpha')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lgt('beta', 'alpha'), "]"
    stop 1
end if
end program t
