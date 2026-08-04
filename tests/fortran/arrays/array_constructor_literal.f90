! vybe-test: fortran/arrays/array_constructor_literal
! origin: languages/fortran/tests/fortran/test_arrays.rs
program t
  integer :: a(3) = [1, 2, 3]
  print *, a(2)
end program t
