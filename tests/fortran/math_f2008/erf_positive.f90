! vybe-test: fortran/math_f2008/erf_positive
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((erf(1.0)) - (0.842700779)) > 8.427008e-06) then
      print *, "FAIL: want [0.842700779] got [", erf(1.0), "]"
      stop 1
  end if
end program t
