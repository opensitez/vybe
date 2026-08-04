! vybe-test: fortran/subroutines/procedure_dummy_array_result_preserves_values
! origin: languages/fortran/tests/fortran/test_subroutines.rs
module m
  implicit none
  abstract interface
    function rhs_func(t, y, n) result(dydt)
      integer, intent(in) :: n
      real, intent(in) :: t
      real, intent(in) :: y(n)
      real :: dydt(n)
    end function rhs_func
  end interface
contains
  subroutine step(rhs)
    procedure(rhs_func) :: rhs
    real :: y(3), dydt(3)
    y = [1.0, 0.0, 0.0]
    dydt = rhs(0.0, y, 3)
    if ((dydt(1)) /= -10) then
    print *, "FAIL: want [-10] got [", dydt(1), "]"
    stop 1
end if
    if ((dydt(2)) /= 28) then
    print *, "FAIL: want [28] got [", dydt(2), "]"
    stop 1
end if
    if ((dydt(3)) /= 0) then
    print *, "FAIL: want [0] got [", dydt(3), "]"
    stop 1
end if
  end subroutine step
end module m

program test
  use m
  call step(lorenz_rhs)
contains
  function lorenz_rhs(t, y, n) result(dydt)
    integer, intent(in) :: n
    real, intent(in) :: t
    real, intent(in) :: y(n)
    real :: dydt(n)
    dydt(1) = 10.0 * (y(2) - y(1))
    dydt(2) = y(1) * (28.0 - y(3)) - y(2)
    dydt(3) = y(1) * y(2)
  end function lorenz_rhs
end program test
