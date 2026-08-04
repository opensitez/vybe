! vybe-test: fortran/arrays/array_dim_attr
! origin: languages/fortran/tests/fortran/test_arrays.rs
program t
  integer, dimension(10) :: a
  a(1) = 99
  print *, a(1)
end program t
