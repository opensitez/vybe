! vybe-test: fortran/where_advanced/bit_size_int1
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer(kind=1) :: x = 0_1
    print *, bit_size(x)
end program test
