! vybe-test: fortran/bits_f2008/merge_bits_basic
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs
program t
  print *, merge_bits(int(z'FF00'), int(z'00FF'), int(z'F0F0'))
end program t
