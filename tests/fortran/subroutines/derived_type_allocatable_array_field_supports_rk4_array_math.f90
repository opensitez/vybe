! vybe-test: fortran/subroutines/derived_type_allocatable_array_field_supports_rk4_array_math
! origin: languages/fortran/tests/fortran/test_subroutines.rs
integer :: vybe_check_i = 0
real :: vybe_check_w(1) = [ 0.01 ]
module m
  implicit none
  type :: ode_state
    real :: t
    real, allocatable :: y(:)
    integer :: neq
  contains
    procedure :: init => state_init
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
  subroutine state_init(self, neq, t0)
    class(ode_state), intent(inout) :: self
    integer, intent(in) :: neq
    real, intent(in) :: t0
    self%neq = neq
    self%t = t0
    allocate(self%y(neq))
    self%y = 0.0
  end subroutine

  subroutine rk4_step(state, h, rhs)
    type(ode_state), intent(inout) :: state
    real, intent(in) :: h
    procedure(rhs_func) :: rhs
    real :: k1(state%neq), k2(state%neq), k3(state%neq), k4(state%neq), h2
    h2 = h * 0.5
    k1 = rhs(state%t, state%y, state%neq)
    k2 = rhs(state%t + h2, state%y + h2 * k1, state%neq)
    k3 = rhs(state%t + h2, state%y + h2 * k2, state%neq)
    k4 = rhs(state%t + h, state%y + h * k3, state%neq)
    state%y = state%y + (h / 6.0) * (k1 + 2.0 * k2 + 2.0 * k3 + k4)
    state%t = state%t + h
  end subroutine
end module m

program test
  use m
  type(ode_state) :: state
  call state%init(3, 0.0)
  state%y = [1.0, 0.0, 0.0]
  call rk4_step(state, 0.01, lorenz_rhs)
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (abs((state%t) - vybe_check_w(vybe_check_i)) > 1.0e-6) then
      print *, "FAIL at ", vybe_check_i, " got [", state%t, "]"
      stop 1
  end if
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (abs((state%y(1)) - vybe_check_w(vybe_check_i)) > 1.0e-6) then
      print *, "FAIL at ", vybe_check_i, " got [", state%y(1), "]"
      stop 1
  end if
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (abs((state%y(2)) - vybe_check_w(vybe_check_i)) > 1.0e-6) then
      print *, "FAIL at ", vybe_check_i, " got [", state%y(2), "]"
      stop 1
  end if
    vybe_check_i = vybe_check_i + 1
  if (vybe_check_i > 1) then
      print *, "FAIL: more than 1 line(s)"
      stop 1
  end if
  if (abs((state%y(3)) - vybe_check_w(vybe_check_i)) > 1.0e-6) then
      print *, "FAIL at ", vybe_check_i, " got [", state%y(3), "]"
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
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
