! vybe-test: fortran/where_advanced/bit_size_int8
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer(kind=8) :: x = 0_8
    print *, bit_size(x)
end program test
