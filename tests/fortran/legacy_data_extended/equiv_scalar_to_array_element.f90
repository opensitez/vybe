! vybe-test: fortran/legacy_data_extended/equiv_scalar_to_array_element
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: arr(3), scalar
    equivalence (arr(2), scalar)
    scalar = 31
    print *, arr(2)
end program t
