! vybe-test: fortran/legacy_data_extended/data_implied_do_nested_2d
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: m(2, 3)
    data m /1, 2, 3, 4, 5, 6/
    print *, m(2, 2)
end program t
