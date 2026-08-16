! vybe-test: fortran/math_f2008/bessel_j0_double
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((bessel_j0(1.0d0)) - (0.7651976865579666)) > 7.651977e-06) then
      print *, "FAIL: want [0.7651976865579666] got [", bessel_j0(1.0d0), "]"
      stop 1
  end if
end program t
