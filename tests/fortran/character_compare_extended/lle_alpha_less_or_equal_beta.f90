! vybe-test: fortran/character_compare_extended/lle_alpha_less_or_equal_beta
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lle('alpha', 'beta')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", lle('alpha', 'beta'), "]"
    stop 1
end if
end program t
