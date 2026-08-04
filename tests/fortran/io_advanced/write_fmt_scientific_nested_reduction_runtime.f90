! vybe-test: fortran/io_advanced/write_fmt_scientific_nested_reduction_runtime
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  integer, parameter :: dp = kind(1.0d0)
  complex(dp) :: spectrum(4), signal(4)
  spectrum = [cmplx(1.0_dp, 0.0_dp, dp), cmplx(2.0_dp, 0.0_dp, dp), cmplx(3.0_dp, 0.0_dp, dp), cmplx(4.0_dp, 0.0_dp, dp)]
  signal = [cmplx(1.0_dp, 0.0_dp, dp), cmplx(2.0_dp, 0.0_dp, dp), cmplx(3.0_dp, 0.0_dp, dp), cmplx(4.001_dp, 0.0_dp, dp)]
  print '(A, ES10.3)', 'delta=', &
      maxval(abs(real(spectrum(1:4), dp) - real(signal(1:4), dp)))
end program t
