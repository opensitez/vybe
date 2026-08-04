! vybe-test: fortran/legacy_data_extended/equiv_read_after_write_overlap
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: alpha, beta
    equivalence (alpha, beta)
    alpha = 77
    print *, beta
end program t
