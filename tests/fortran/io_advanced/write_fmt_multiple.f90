! vybe-test: fortran/io_advanced/write_fmt_multiple
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  integer :: n = 42
  real :: x = 3.14
  write(*, '(I5, F8.3)') n, x
end program t
