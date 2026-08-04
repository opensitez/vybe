! vybe-test: fortran/legacy_data_extended/equiv_chain_three_names
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: p, q, r
    equivalence (p, q, r)
    p = 12
    print *, q, r
end program t
