! vybe-test: fortran/bit_btest_ishftc/ishftc_assign_then_btest_result_bit
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs

program t
    integer :: x
    x = ishftc(5, 2, 4)
    print *, x
    print *, btest(x, 2)
end program t
