! vybe-test: fortran/character_compare_extended/lge_not_less_than_equal
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
if ((lge('alpha', 'beta')) .neqv. .false.) then
    print *, "FAIL: want [false] got [", lge('alpha', 'beta'), "]"
    stop 1
end if
end program t
