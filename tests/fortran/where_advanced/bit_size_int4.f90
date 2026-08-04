! vybe-test: fortran/where_advanced/bit_size_int4
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: x = 0
    print *, bit_size(x)
end program test
