! vybe-test: fortran/math_f2008/gamma_five
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((gamma(5.0)) - (24.0)) > 2.400000e-04) then
      print *, "FAIL: want [24.0] got [", gamma(5.0), "]"
      stop 1
  end if
end program t
