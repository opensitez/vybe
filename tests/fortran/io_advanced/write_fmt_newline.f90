! vybe-test: fortran/io_advanced/write_fmt_newline
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  write(*, '(A, /, A)') 'line1', 'line2'
end program t
