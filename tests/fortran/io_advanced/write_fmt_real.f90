! vybe-test: fortran/io_advanced/write_fmt_real
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  write(*, '(F8.3)') 3.14159
end program t
