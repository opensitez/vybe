! vybe-test: fortran/math_f2008/erfc_scaled_basic
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  if (abs((erfc_scaled(1.0)) - (0.427583575)) > 4.275836e-06) then
      print *, "FAIL: want [0.427583575] got [", erfc_scaled(1.0), "]"
      stop 1
  end if
end program t
