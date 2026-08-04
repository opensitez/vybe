! vybe-test: fortran/bit_btest_ishftc/btest_negative_one_sign_bit
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
  print *, btest(-1, 31)
end program t
