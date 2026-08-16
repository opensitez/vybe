! vybe-test: fortran/math_f2008/gamma_one
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((gamma(1.0)) - (1.0)) > 1.000000e-05) then
      print *, "FAIL: want [1.0] got [", gamma(1.0), "]"
      stop 1
  end if
end program t
