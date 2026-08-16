! vybe-test: fortran/math_f2008/erf_zero
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((erf(0.0)) - (0.0)) > 1.000000e-06) then
      print *, "FAIL: want [0.0] got [", erf(0.0), "]"
      stop 1
  end if
end program t
