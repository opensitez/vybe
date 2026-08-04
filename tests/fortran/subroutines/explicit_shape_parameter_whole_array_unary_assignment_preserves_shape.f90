! vybe-test: fortran/subroutines/explicit_shape_parameter_whole_array_unary_assignment_preserves_shape
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program test
  real :: a(2,2), b(2,2)
  call foo(a, b)
  if ((sum(a)) /= 16) then
    print *, "FAIL: want [16] got [", sum(a), "]"
    stop 1
end if
  if ((sum(b)) /= -16) then
    print *, "FAIL: want [-16] got [", sum(b), "]"
    stop 1
end if
contains
  subroutine foo(a, b)
    real, intent(out) :: a(2,2)
    real, intent(out) :: b(2,2)
    a = 4.0
    b = -a
  end subroutine foo
end program test
