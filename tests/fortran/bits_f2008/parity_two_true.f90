! vybe-test: fortran/bits_f2008/parity_two_true
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs
program t
  logical :: a(2) = [.true., .true.]
  print *, parity(a)
end program t
