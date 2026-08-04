! vybe-test: fortran/bit_btest_ishftc/ishftc_with_variable_size_operand
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs

program t
    integer :: n = 4
    print *, ishftc(6, 1, n)
end program t
