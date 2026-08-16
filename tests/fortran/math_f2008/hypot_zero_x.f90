! vybe-test: fortran/math_f2008/hypot_zero_x
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((hypot(0.0, 5.0)) - (5.0)) > 5.000000e-05) then
      print *, "FAIL: want [5.0] got [", hypot(0.0, 5.0), "]"
      stop 1
  end if
end program t
