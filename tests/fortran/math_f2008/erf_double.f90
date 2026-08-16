! vybe-test: fortran/math_f2008/erf_double
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((erf(1.0d0)) - (0.8427007929497149)) > 8.427008e-06) then
      print *, "FAIL: want [0.8427007929497149] got [", erf(1.0d0), "]"
      stop 1
  end if
end program t
