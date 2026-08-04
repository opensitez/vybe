! vybe-test: fortran/bit_btest_ishftc/ishftc_zero_shift_preserves_field
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
  print *, ishftc(170, 0, 8)
end program t
