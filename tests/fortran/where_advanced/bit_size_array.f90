! vybe-test: fortran/where_advanced/bit_size_array
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: a(10)
    print *, bit_size(a(1))
end program test
