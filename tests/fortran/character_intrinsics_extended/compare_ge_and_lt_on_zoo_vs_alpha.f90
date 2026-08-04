! vybe-test: fortran/character_intrinsics_extended/compare_ge_and_lt_on_zoo_vs_alpha
! origin: languages/fortran/tests/fortran/test_character_intrinsics_extended.rs
program t
if (('zoo' >= 'alpha') .neqv. .true.) then
    print *, "FAIL: want [true] got [", 'zoo' >= 'alpha', "]"
    stop 1
end if
if (('zoo' < 'alpha') .neqv. .false.) then
    print *, "FAIL: want [false] got [", 'zoo' < 'alpha', "]"
    stop 1
end if
end program t
