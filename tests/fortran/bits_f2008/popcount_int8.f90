! vybe-test: fortran/bits_f2008/popcount_int8
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs
program t
  integer(kind=8) :: x = 1152921504606846975_8
  print *, popcnt(x)
end program t
