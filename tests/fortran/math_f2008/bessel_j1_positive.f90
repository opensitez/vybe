! vybe-test: fortran/math_f2008/bessel_j1_positive
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((bessel_j1(1.0)) - (0.440050572)) > 4.400506e-06) then
      print *, "FAIL: want [0.440050572] got [", bessel_j1(1.0), "]"
      stop 1
  end if
end program t
