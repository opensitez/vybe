! vybe-test: fortran/io_advanced/write_fmt_repeat
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  write(*, '(3I4)') 1, 2, 3
end program t
