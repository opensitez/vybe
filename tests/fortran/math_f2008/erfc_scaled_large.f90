! vybe-test: fortran/math_f2008/erfc_scaled_large
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((erfc_scaled(10.0)) - (0.0561409928)) > 1.000000e-06) then
      print *, "FAIL: want [0.0561409928] got [", erfc_scaled(10.0), "]"
      stop 1
  end if
end program t
