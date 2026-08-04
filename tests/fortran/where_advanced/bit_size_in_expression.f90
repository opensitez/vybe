! vybe-test: fortran/where_advanced/bit_size_in_expression
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: x = 1
    integer :: half_bits
    half_bits = bit_size(x) / 2
    print *, half_bits
end program test
