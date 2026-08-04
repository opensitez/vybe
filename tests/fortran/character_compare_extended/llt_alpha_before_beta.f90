! vybe-test: fortran/character_compare_extended/llt_alpha_before_beta
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((llt('alpha', 'beta')) .neqv. .true.) then
    print *, "FAIL: want [true] got [", llt('alpha', 'beta'), "]"
    stop 1
end if
end program t
