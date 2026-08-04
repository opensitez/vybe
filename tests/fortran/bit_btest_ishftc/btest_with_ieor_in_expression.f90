! vybe-test: fortran/bit_btest_ishftc/btest_with_ieor_in_expression
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs

program t
    integer :: x = 42
    print *, btest(ieor(x, ishft(1, 1)), 1)
end program t
