! vybe-test: fortran/io_advanced/write_fmt_octal
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  write(*, '(O8)') 255
end program t
