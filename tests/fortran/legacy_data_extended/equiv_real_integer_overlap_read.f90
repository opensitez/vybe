! vybe-test: fortran/legacy_data_extended/equiv_real_integer_overlap_read
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    real :: x
    integer :: k
    equivalence (x, k)
    x = 2.0
    print *, k
end program t
