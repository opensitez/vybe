! vybe-test: fortran/where_advanced/storage_size_integer
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: x = 0
    print *, storage_size(x)
end program test
