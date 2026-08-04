! vybe-test: fortran/io_advanced/read_formatted
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  integer :: n
  read(*, '(I5)') n
  print *, n
end program t
