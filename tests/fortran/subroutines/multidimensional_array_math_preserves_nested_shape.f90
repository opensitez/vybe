! vybe-test: fortran/subroutines/multidimensional_array_math_preserves_nested_shape
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program test
  real, allocatable :: a(:,:), b(:,:), c(:,:)
  allocate(a(2,2), b(2,2), c(2,2))
  a = 1.0
  b = 2.0
  c = a + b
  if ((c(1,1)) /= 3) then
    print *, "FAIL: want [3] got [", c(1,1), "]"
    stop 1
end if
  if ((c(2,2)) /= 3) then
    print *, "FAIL: want [3] got [", c(2,2), "]"
    stop 1
end if
  if ((sum(c)) /= 12) then
    print *, "FAIL: want [12] got [", sum(c), "]"
    stop 1
end if
end program test
