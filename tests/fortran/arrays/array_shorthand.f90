! vybe-test: fortran/arrays/array_shorthand
! origin: languages/fortran/tests/fortran/test_arrays.rs
program t
  real :: v(3)
  v(1) = 1.0
  v(2) = 2.0
  v(3) = 3.0
  print *, v(2)
end program t
