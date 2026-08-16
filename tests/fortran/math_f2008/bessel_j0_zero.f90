! vybe-test: fortran/math_f2008/bessel_j0_zero
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((bessel_j0(0.0)) - (1.0)) > 1.000000e-05) then
      print *, "FAIL: want [1.0] got [", bessel_j0(0.0), "]"
      stop 1
  end if
end program t
