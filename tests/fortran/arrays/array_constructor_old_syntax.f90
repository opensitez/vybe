! vybe-test: fortran/arrays/array_constructor_old_syntax
! origin: languages/fortran/tests/fortran/test_arrays.rs
program t
  integer :: a(3) = (/1, 2, 3/)
  print *, a(1)
end program t
