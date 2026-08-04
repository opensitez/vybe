! vybe-test: fortran/legacy_data_extended/equiv_multiple_independent_groups
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: a, b, c, d
    equivalence (a, b)
    equivalence (c, d)
    a = 1
    c = 9
    print *, b, d
end program t
