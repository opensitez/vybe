! vybe-test: fortran/math_f2008/hypot_zero_x
! origin: languages/fortran/tests/fortran/test_math_f2008.rs
program t
  print *, hypot(0.0, 5.0)
end program t
