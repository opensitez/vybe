! vybe-test: fortran/math_f2008/erfc_one
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((erfc(1.0)) - (0.157299206)) > 1.572992e-06) then
      print *, "FAIL: want [0.157299206] got [", erfc(1.0), "]"
      stop 1
  end if
end program t
