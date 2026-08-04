! vybe-test: fortran/io_advanced/read_list_directed
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  integer :: n
  read(*, *) n
  print *, n
end program t
