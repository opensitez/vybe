! vybe-test: fortran/io_advanced/write_fmt_double
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  write(*, '(D20.12)') 3.141592653589793
end program t
