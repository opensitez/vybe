! vybe-test: fortran/bits_f2008/dshiftl_zero_shift
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs
program t
  print *, dshiftl(42, 0, 0)
end program t
