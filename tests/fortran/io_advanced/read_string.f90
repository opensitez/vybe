! vybe-test: fortran/io_advanced/read_string
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  character(len=20) :: s
  read(*, *) s
  print *, s
end program t
