! vybe-test: fortran/where_advanced/bit_size_vs_storage_size
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: x = 0
    print *, bit_size(x) == storage_size(x)
end program test
