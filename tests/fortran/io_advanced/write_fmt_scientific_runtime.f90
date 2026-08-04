! vybe-test: fortran/io_advanced/write_fmt_scientific_runtime
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  integer, parameter :: dp = kind(1.0d0)
  print '(ES14.4)', 0.25_dp
  print '(A, ES10.3)', 'value=', 0.25_dp
end program t
