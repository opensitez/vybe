! vybe-test: fortran/math_f2008/gamma_half
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((gamma(0.5)) - (1.7724539)) > 1.772454e-05) then
      print *, "FAIL: want [1.7724539] got [", gamma(0.5), "]"
      stop 1
  end if
end program t
