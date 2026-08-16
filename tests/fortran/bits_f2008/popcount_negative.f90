! vybe-test: fortran/bits_f2008/popcount_negative
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs
program t
  integer :: x = -1
  print *, popcnt(x)
end program t
