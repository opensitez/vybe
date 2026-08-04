! vybe-test: fortran/bits_f2008/leadz_kind8
! origin: languages/fortran/tests/fortran/test_bits_f2008.rs
program t
  integer(kind=8) :: x = 1_8
  print *, leadz(x)
end program t
