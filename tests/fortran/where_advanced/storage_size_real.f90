! vybe-test: fortran/where_advanced/storage_size_real
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    real :: x = 0.0
    print *, storage_size(x)
end program test
