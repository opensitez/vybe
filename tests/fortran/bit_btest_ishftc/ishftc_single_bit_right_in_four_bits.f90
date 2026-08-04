! vybe-test: fortran/bit_btest_ishftc/ishftc_single_bit_right_in_four_bits
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
  print *, ishftc(8, -1, 4)
end program t
