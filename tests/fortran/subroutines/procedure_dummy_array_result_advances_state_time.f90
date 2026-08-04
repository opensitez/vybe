! vybe-test: fortran/subroutines/procedure_dummy_array_result_advances_state_time
! origin: languages/fortran/tests/fortran/test_subroutines.rs
module m
  implicit none
  type :: ode_state
    real :: t
    real :: y(3)
  end type ode_state
  abstract interface
    function rhs_func(t, y, n) result(dydt)
      integer, intent(in) :: n
      real, intent(in) :: t
      real, intent(in) :: y(n)
      real :: dydt(n)
    end function rhs_func
  end interface
contains
  subroutine rk4_step(state, h, rhs)
    type(ode_state), intent(inout) :: state
    real, intent(in) :: h
    procedure(rhs_func) :: rhs
    real :: k1(3)
    k1 = rhs(state%t, state%y, 3)
    state%y = state%y + h * k1
    state%t = state%t + h
  end subroutine rk4_step
end module m

program test
  use m
  type(ode_state) :: state
  state%t = 0.0
  state%y = [1.0, 0.0, 0.0]
  call rk4_step(state, 0.5, lorenz_rhs)
  if (abs((state%t) - 0.5) > 1.0e-6) then
    print *, "FAIL: want [0.5] got [", state%t, "]"
    stop 1
end if
  if ((state%y(1)) /= -4) then
    print *, "FAIL: want [-4] got [", state%y(1), "]"
    stop 1
end if
  if ((state%y(2)) /= 14) then
    print *, "FAIL: want [14] got [", state%y(2), "]"
    stop 1
end if
  if ((state%y(3)) /= 0) then
    print *, "FAIL: want [0] got [", state%y(3), "]"
    stop 1
end if
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
