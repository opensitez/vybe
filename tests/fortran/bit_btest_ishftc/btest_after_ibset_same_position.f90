! vybe-test: fortran/bit_btest_ishftc/btest_after_ibset_same_position
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs

program t
    integer :: x
    x = ibset(0, 7)
    print *, btest(x, 7)
    print *, btest(x, 6)
end program t
