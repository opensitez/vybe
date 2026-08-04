! vybe-test: fortran/bits_f2008/leadz_max_int
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs
program t
  print *, leadz(huge(0))
end program t
