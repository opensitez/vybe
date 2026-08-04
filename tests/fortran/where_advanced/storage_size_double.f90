! vybe-test: fortran/where_advanced/storage_size_double
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    real(kind=8) :: x = 0.0d0
    print *, storage_size(x)
end program test
