! vybe-test: fortran/math_f2008/asinh_zero
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((asinh(0.0)) - (0.0)) > 1.000000e-06) then
      print *, "FAIL: want [0.0] got [", asinh(0.0), "]"
      stop 1
  end if
end program t
