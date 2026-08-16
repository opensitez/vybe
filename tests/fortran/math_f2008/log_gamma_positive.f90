! vybe-test: fortran/math_f2008/log_gamma_positive
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((log_gamma(5.0)) - (3.17805386)) > 3.178054e-05) then
      print *, "FAIL: want [3.17805386] got [", log_gamma(5.0), "]"
      stop 1
  end if
end program t
