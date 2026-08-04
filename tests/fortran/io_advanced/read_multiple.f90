! vybe-test: fortran/io_advanced/read_multiple
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  integer :: a, b
  read(*, *) a, b
  print *, a + b
end program t
