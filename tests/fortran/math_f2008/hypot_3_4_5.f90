! vybe-test: fortran/math_f2008/hypot_3_4_5
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((hypot(3.0, 4.0)) - (5.0)) > 5.000000e-05) then
      print *, "FAIL: want [5.0] got [", hypot(3.0, 4.0), "]"
      stop 1
  end if
end program t
