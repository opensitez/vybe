! vybe-test: fortran/io_advanced/write_fmt_logical
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  write(*, '(L5)') .true.
end program t
