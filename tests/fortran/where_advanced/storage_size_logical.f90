! vybe-test: fortran/where_advanced/storage_size_logical
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    logical :: b = .false.
    print *, storage_size(b)
end program test
