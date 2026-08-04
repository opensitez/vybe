! vybe-test: fortran/arrays/array_constructor_implied_do
! origin: languages/fortran/tests/fortran/test_arrays.rs
program t
  integer :: a(5) = [(i, i=1,5)]
  print *, a(3)
end program t
