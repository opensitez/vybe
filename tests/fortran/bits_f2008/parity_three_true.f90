! vybe-test: fortran/bits_f2008/parity_three_true
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs
program t
  logical :: a(3) = [.true., .true., .true.]
  print *, parity(a)
end program t
