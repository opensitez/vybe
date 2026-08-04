! vybe-test: fortran/io_advanced/write_fmt_scientific
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  write(*, '(E12.4)') 1.23456e10
end program t
