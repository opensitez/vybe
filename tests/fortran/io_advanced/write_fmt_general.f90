! vybe-test: fortran/io_advanced/write_fmt_general
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  write(*, '(G12.4)') 3.14
end program t
