! vybe-test: fortran/io_advanced/write_fmt_hex
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  write(*, '(Z8)') 255
end program t
