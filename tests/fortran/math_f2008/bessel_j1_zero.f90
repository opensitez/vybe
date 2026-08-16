! vybe-test: fortran/math_f2008/bessel_j1_zero
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((bessel_j1(0.0)) - (0.0)) > 1.000000e-06) then
      print *, "FAIL: want [0.0] got [", bessel_j1(0.0), "]"
      stop 1
  end if
end program t
