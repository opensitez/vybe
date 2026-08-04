! vybe-test: fortran/legacy_data_extended/equiv_adjacent_array_elements
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: seq(4)
    equivalence (seq(2), seq(3))
    seq(2) = 55
    print *, seq(3)
end program t
