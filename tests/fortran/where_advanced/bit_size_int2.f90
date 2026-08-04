! vybe-test: fortran/where_advanced/bit_size_int2
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer(kind=2) :: x = 0_2
    print *, bit_size(x)
end program test
