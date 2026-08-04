! vybe-test: fortran/complex/complex_slice_real_kind_abs_maxval_runtime
! origin: languages/fortran/tests/fortran/test_complex.rs
program t
  integer, parameter :: dp = kind(1.0d0)
  complex(dp) :: a(4), b(4)
  a(1) = cmplx(1.0_dp, 0.0_dp, dp)
  a(2) = cmplx(0.5_dp, 0.0_dp, dp)
  a(3) = cmplx(0.0_dp, 0.0_dp, dp)
  a(4) = cmplx(0.0_dp, 0.0_dp, dp)
  b = a
  if ((maxval(abs(real(a(1:4), dp) - real(b(1:4), dp)))) /= 0) then
    print *, "FAIL: want [0] got [", maxval(abs(real(a(1:4), dp) - real(b(1:4), dp))), "]"
    stop 1
end if
end program t
