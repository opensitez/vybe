! vybe-test: fortran/math_f2008/bessel_j0_positive
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((bessel_j0(1.0)) - (0.765197694)) > 7.651977e-06) then
      print *, "FAIL: want [0.765197694] got [", bessel_j0(1.0), "]"
      stop 1
  end if
end program t
