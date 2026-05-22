! ODE Solver — 4th-order Runge-Kutta integration
! Covers: derived types, procedure pointers, modules, allocatable arrays,
!         recursive functions, select case, where construct.
!
! Solves the Lorenz attractor system:
!   dx/dt = sigma*(y - x)
!   dy/dt = x*(rho - z) - y
!   dz/dt = x*y - beta*z

module ode_module
    implicit none

    integer, parameter :: dp = kind(1.0d0)

    ! Derived type for ODE state
    type :: ode_state
        real(dp) :: t
        real(dp), allocatable :: y(:)
        integer  :: neq
    contains
        procedure :: init   => state_init
        procedure :: copy   => state_copy
        procedure :: norm   => state_norm
    end type ode_state

    ! Abstract interface for RHS function
    abstract interface
        function rhs_func(t, y, n) result(dydt)
            import dp
            integer,  intent(in) :: n
            real(dp), intent(in) :: t
            real(dp), intent(in) :: y(n)
            real(dp) :: dydt(n)
        end function rhs_func
    end interface

contains

    subroutine state_init(self, neq, t0)
        class(ode_state), intent(inout) :: self
        integer,  intent(in) :: neq
        real(dp), intent(in) :: t0
        self%neq = neq
        self%t   = t0
        allocate(self%y(neq))
        self%y = 0.0_dp
    end subroutine state_init

    subroutine state_copy(self, other)
        class(ode_state), intent(inout) :: self
        type(ode_state),  intent(in)    :: other
        self%t   = other%t
        self%neq = other%neq
        if (allocated(self%y)) deallocate(self%y)
        allocate(self%y(other%neq))
        self%y = other%y
    end subroutine state_copy

    pure function state_norm(self) result(n)
        class(ode_state), intent(in) :: self
        real(dp) :: n
        n = sqrt(sum(self%y**2))
    end function state_norm

    ! 4th-order Runge-Kutta step
    subroutine rk4_step(state, h, rhs)
        type(ode_state), intent(inout) :: state
        real(dp),        intent(in)    :: h
        procedure(rhs_func)            :: rhs

        real(dp), dimension(state%neq) :: k1, k2, k3, k4
        real(dp) :: t, h2

        t  = state%t
        h2 = h * 0.5_dp

        k1 = rhs(t,      state%y,            state%neq)
        k2 = rhs(t + h2, state%y + h2 * k1,  state%neq)
        k3 = rhs(t + h2, state%y + h2 * k2,  state%neq)
        k4 = rhs(t + h,  state%y + h  * k3,  state%neq)

        state%y = state%y + (h / 6.0_dp) * (k1 + 2.0_dp*k2 + 2.0_dp*k3 + k4)
        state%t = t + h
    end subroutine rk4_step

    ! Adaptive RK4 with error control
    subroutine rk4_adaptive(state, h, tol, rhs, h_new)
        type(ode_state), intent(inout) :: state
        real(dp),        intent(inout) :: h
        real(dp),        intent(in)    :: tol
        procedure(rhs_func)            :: rhs
        real(dp),        intent(out)   :: h_new

        type(ode_state) :: s1, s2
        real(dp) :: err, scale

        call s1%copy(state)
        call s2%copy(state)

        ! Full step
        call rk4_step(s1, h, rhs)

        ! Two half steps
        call rk4_step(s2, h * 0.5_dp, rhs)
        call rk4_step(s2, h * 0.5_dp, rhs)

        ! Error estimate
        err = sqrt(sum((s1%y - s2%y)**2)) / (15.0_dp * max(s2%norm(), 1.0e-10_dp))

        if (err < tol) then
            call state%copy(s2)  ! accept the more accurate two-step result
            scale = (tol / max(err, 1.0e-15_dp))**0.2_dp
            h_new = min(h * min(scale * 0.9_dp, 5.0_dp), 0.1_dp)
        else
            ! Reject step, reduce h
            scale = (tol / max(err, 1.0e-15_dp))**0.25_dp
            h_new = h * max(scale * 0.9_dp, 0.1_dp)
        end if
    end subroutine rk4_adaptive

end module ode_module


! Lorenz attractor RHS (module-level, not contained — demonstrates external procedure)
function lorenz_rhs(t, y, n) result(dydt)
    use ode_module, only: dp
    implicit none
    integer,  intent(in) :: n
    real(dp), intent(in) :: t, y(n)
    real(dp) :: dydt(n)

    real(dp), parameter :: sigma = 10.0_dp
    real(dp), parameter :: rho   = 28.0_dp
    real(dp), parameter :: beta  = 8.0_dp / 3.0_dp

    dydt(1) = sigma * (y(2) - y(1))
    dydt(2) = y(1) * (rho - y(3)) - y(2)
    dydt(3) = y(1) * y(2) - beta * y(3)
end function lorenz_rhs


program ode_solver
    use ode_module
    implicit none

    interface
        function lorenz_rhs(t, y, n) result(dydt)
            import dp
            integer,  intent(in) :: n
            real(dp), intent(in) :: t, y(n)
            real(dp) :: dydt(n)
        end function lorenz_rhs
    end interface

    type(ode_state) :: state
    real(dp) :: h, h_new, t_end, tol
    integer  :: step, max_steps, print_every
    real(dp), allocatable :: trajectory(:, :)
    integer :: n_saved, i

    ! Initial conditions
    call state%init(3, 0.0_dp)
    state%y = [1.0_dp, 0.0_dp, 0.0_dp]

    h          = 0.01_dp
    t_end      = 50.0_dp
    tol        = 1.0e-6_dp
    max_steps  = 1000000
    print_every = 500

    ! Allocate trajectory storage
    allocate(trajectory(4, max_steps / print_every + 1))
    n_saved = 0

    print *, "Integrating Lorenz attractor..."
    print *, "  sigma=10, rho=28, beta=8/3"
    print *, "  IC: (1, 0, 0),  t=[0, 50]"
    print *, ""
    print "(a6, 4a12)", "Step", "t", "x", "y", "z"
    print "(a6, 4a12)", "----", "-", "-", "-", "-"

    do step = 1, max_steps
        call rk4_adaptive(state, h, tol, lorenz_rhs, h_new)
        h = h_new

        if (mod(step, print_every) == 0) then
            print "(i6, 4f12.4)", step, state%t, state%y(1), state%y(2), state%y(3)
            n_saved = n_saved + 1
            trajectory(1, n_saved) = state%t
            trajectory(2:4, n_saved) = state%y
        end if

        if (state%t >= t_end) exit
    end do

    print *, ""
    print "(a, i0, a)", "Completed ", step, " steps"
    print "(a, f8.4)", "Final t = ", state%t
    print "(a, f8.4)", "Final |y| = ", state%norm()

    ! Statistics on trajectory
    print *, ""
    print *, "=== Trajectory statistics ==="
    do i = 1, 3
        print "(a, i0, a, 2f10.4)", "  Component ", i, &
            ":  min/max = ", &
            minval(trajectory(i+1, 1:n_saved)), &
            maxval(trajectory(i+1, 1:n_saved))
    end do

    deallocate(trajectory)

end program ode_solver
