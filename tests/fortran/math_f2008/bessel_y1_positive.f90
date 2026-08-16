! vybe-test: fortran/math_f2008/bessel_y1_positive
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((bessel_y1(1.0)) - (-0.781212807)) > 7.812128e-06) then
      print *, "FAIL: want [-0.781212807] got [", bessel_y1(1.0), "]"
      stop 1
  end if
end program t
