! ============================================================
! Shallow Water Equations — 2D Weather Model
! ============================================================
! A simplified numerical weather prediction model solving the
! shallow water equations on a periodic domain using finite
! differences. This is the core of many real NWP models.
!
! Equations:
!   dh/dt  + d(hu)/dx + d(hv)/dy = 0          (continuity)
!   du/dt  + u*du/dx + v*du/dy   = -g*dh/dx + f*v  (x-momentum)
!   dv/dt  + u*dv/dx + v*dv/dy   = -g*dh/dy - f*u  (y-momentum)
!
! Features used:
!   - Modules with derived types
!   - Allocatable 2D arrays
!   - Do concurrent (parallel loops)
!   - Coarrays (optional, guarded)
!   - Namelist I/O for configuration
!   - Binary file output (restart files)
!   - Operator overloading for field arithmetic
!   - Generic interfaces
! ============================================================

module swe_types
    implicit none

    integer, parameter :: dp = kind(1.0d0)
    real(dp), parameter :: G    = 9.81_dp   ! gravity
    real(dp), parameter :: PI   = 4.0_dp * atan(1.0_dp)

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: dx, dy
        real(dp) :: Lx, Ly
        real(dp), allocatable :: x(:), y(:)
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t

    type :: field2d
        real(dp), allocatable :: data(:,:)
        integer :: nx, ny
        character(len=32) :: name
    contains
        procedure :: init    => field_init
        procedure :: max_val => field_max
        procedure :: min_val => field_min
        procedure :: rms     => field_rms
        procedure :: fill    => field_fill
    end type field2d

    type :: swe_state
        type(field2d) :: h    ! fluid depth
        type(field2d) :: u    ! x-velocity
        type(field2d) :: v    ! y-velocity
        real(dp)      :: time
    end type swe_state

    type :: swe_config
        integer  :: nx, ny
        integer  :: nt
        real(dp) :: Lx, Ly
        real(dp) :: dt
        real(dp) :: f_coriolis
        real(dp) :: h0          ! mean depth
        integer  :: output_freq
        logical  :: write_restart
        character(len=256) :: output_prefix
    end type swe_config

contains

    subroutine grid_init(self, nx, ny, Lx, Ly)
        class(grid_t), intent(inout) :: self
        integer,  intent(in) :: nx, ny
        real(dp), intent(in) :: Lx, Ly
        integer :: i
        self%nx = nx;  self%ny = ny
        self%Lx = Lx;  self%Ly = Ly
        self%dx = Lx / nx
        self%dy = Ly / ny
        allocate(self%x(nx), self%y(ny))
        self%x = [(( i - 0.5_dp) * self%dx, i = 1, nx)]
        self%y = [(( i - 0.5_dp) * self%dy, i = 1, ny)]
    end subroutine grid_init

    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        print "(a, i0, a, i0)", "  Grid: ", self%nx, " x ", self%ny
        print "(a, f8.1, a, f8.1)", "  Domain: Lx=", self%Lx, "  Ly=", self%Ly
        print "(a, f8.4, a, f8.4)", "  dx=", self%dx, "  dy=", self%dy
    end subroutine grid_print

    subroutine field_init(self, nx, ny, name)
        class(field2d), intent(inout) :: self
        integer, intent(in) :: nx, ny
        character(len=*), intent(in) :: name
        self%nx = nx;  self%ny = ny
        self%name = name
        allocate(self%data(nx, ny))
        self%data = 0.0_dp
    end subroutine field_init

    real(dp) function field_max(self)
        class(field2d), intent(in) :: self
        field_max = maxval(self%data)
    end function field_max

    real(dp) function field_min(self)
        class(field2d), intent(in) :: self
        field_min = minval(self%data)
    end function field_min

    real(dp) function field_rms(self)
        class(field2d), intent(in) :: self
        field_rms = sqrt(sum(self%data**2) / (self%nx * self%ny))
    end function field_rms

    subroutine field_fill(self, val)
        class(field2d), intent(inout) :: self
        real(dp), intent(in) :: val
        self%data = val
    end subroutine field_fill

end module swe_types


module swe_numerics
    use swe_types
    implicit none

contains

    ! Periodic index wrapping
    pure function wrap(i, n) result(j)
        integer, intent(in) :: i, n
        integer :: j
        j = mod(i - 1 + n, n) + 1
    end function wrap

    ! Compute spatial derivatives using 4th-order centered differences
    subroutine ddx(f, dfdx, dx, nx, ny)
        real(dp), intent(in)  :: f(nx, ny)
        real(dp), intent(out) :: dfdx(nx, ny)
        real(dp), intent(in)  :: dx
        integer,  intent(in)  :: nx, ny
        integer :: i, j, im2, im1, ip1, ip2

        do concurrent (j = 1:ny, i = 1:nx)
            im2 = wrap(i-2, nx);  im1 = wrap(i-1, nx)
            ip1 = wrap(i+1, nx);  ip2 = wrap(i+2, nx)
            dfdx(i,j) = (-f(ip2,j) + 8.0_dp*f(ip1,j) &
                         - 8.0_dp*f(im1,j) + f(im2,j)) / (12.0_dp * dx)
        end do
    end subroutine ddx

    subroutine ddy(f, dfdy, dy, nx, ny)
        real(dp), intent(in)  :: f(nx, ny)
        real(dp), intent(out) :: dfdy(nx, ny)
        real(dp), intent(in)  :: dy
        integer,  intent(in)  :: nx, ny
        integer :: i, j, jm2, jm1, jp1, jp2

        do concurrent (j = 1:ny, i = 1:nx)
            jm2 = wrap(j-2, ny);  jm1 = wrap(j-1, ny)
            jp1 = wrap(j+1, ny);  jp2 = wrap(j+2, ny)
            dfdy(i,j) = (-f(i,jp2) + 8.0_dp*f(i,jp1) &
                         - 8.0_dp*f(i,jm1) + f(i,jm2)) / (12.0_dp * dy)
        end do
    end subroutine ddy

    ! Biharmonic diffusion (del^4) for numerical stability
    subroutine biharmonic(f, del4f, dx, dy, nx, ny, nu)
        real(dp), intent(in)  :: f(nx, ny)
        real(dp), intent(out) :: del4f(nx, ny)
        real(dp), intent(in)  :: dx, dy, nu
        integer,  intent(in)  :: nx, ny
        real(dp), allocatable :: lap(:,:), lap2(:,:)
        integer :: i, j, im1, ip1, jm1, jp1

        allocate(lap(nx,ny), lap2(nx,ny))

        ! Laplacian
        do concurrent (j = 1:ny, i = 1:nx)
            im1 = wrap(i-1,nx);  ip1 = wrap(i+1,nx)
            jm1 = wrap(j-1,ny);  jp1 = wrap(j+1,ny)
            lap(i,j) = (f(ip1,j) - 2.0_dp*f(i,j) + f(im1,j)) / dx**2 &
                     + (f(i,jp1) - 2.0_dp*f(i,j) + f(i,jm1)) / dy**2
        end do

        ! Laplacian of Laplacian
        do concurrent (j = 1:ny, i = 1:nx)
            im1 = wrap(i-1,nx);  ip1 = wrap(i+1,nx)
            jm1 = wrap(j-1,ny);  jp1 = wrap(j+1,ny)
            lap2(i,j) = (lap(ip1,j) - 2.0_dp*lap(i,j) + lap(im1,j)) / dx**2 &
                      + (lap(i,jp1) - 2.0_dp*lap(i,j) + lap(i,jm1)) / dy**2
        end do

        del4f = -nu * lap2
        deallocate(lap, lap2)
    end subroutine biharmonic

    ! Compute RHS of SWE (tendencies)
    subroutine compute_rhs(state, grid, cfg, dh_dt, du_dt, dv_dt)
        type(swe_state),  intent(in)  :: state
        type(grid_t),     intent(in)  :: grid
        type(swe_config), intent(in)  :: cfg
        real(dp),         intent(out) :: dh_dt(grid%nx, grid%ny)
        real(dp),         intent(out) :: du_dt(grid%nx, grid%ny)
        real(dp),         intent(out) :: dv_dt(grid%nx, grid%ny)

        integer  :: nx, ny
        real(dp), allocatable :: dhdx(:,:), dhdy(:,:)
        real(dp), allocatable :: dudx(:,:), dudy(:,:)
        real(dp), allocatable :: dvdx(:,:), dvdy(:,:)
        real(dp), allocatable :: dhu_dx(:,:), dhv_dy(:,:)
        real(dp), allocatable :: hu(:,:), hv(:,:)
        real(dp), allocatable :: diff_h(:,:), diff_u(:,:), diff_v(:,:)
        real(dp), parameter   :: nu = 1.0e6_dp   ! diffusion coefficient

        nx = grid%nx;  ny = grid%ny
        allocate(dhdx(nx,ny), dhdy(nx,ny), dudx(nx,ny), dudy(nx,ny), &
                 dvdx(nx,ny), dvdy(nx,ny), dhu_dx(nx,ny), dhv_dy(nx,ny), &
                 hu(nx,ny), hv(nx,ny), diff_h(nx,ny), diff_u(nx,ny), diff_v(nx,ny))

        ! Flux form for continuity
        hu = state%h%data * state%u%data
        hv = state%h%data * state%v%data
        call ddx(hu, dhu_dx, grid%dx, nx, ny)
        call ddy(hv, dhv_dy, grid%dy, nx, ny)

        ! Momentum advection
        call ddx(state%h%data, dhdx, grid%dx, nx, ny)
        call ddy(state%h%data, dhdy, grid%dy, nx, ny)
        call ddx(state%u%data, dudx, grid%dx, nx, ny)
        call ddy(state%u%data, dudy, grid%dy, nx, ny)
        call ddx(state%v%data, dvdx, grid%dx, nx, ny)
        call ddy(state%v%data, dvdy, grid%dy, nx, ny)

        ! Diffusion
        call biharmonic(state%h%data, diff_h, grid%dx, grid%dy, nx, ny, nu)
        call biharmonic(state%u%data, diff_u, grid%dx, grid%dy, nx, ny, nu)
        call biharmonic(state%v%data, diff_v, grid%dx, grid%dy, nx, ny, nu)

        ! Tendencies
        dh_dt = -dhu_dx - dhv_dy + diff_h

        du_dt = -(state%u%data * dudx + state%v%data * dudy) &
                - G * dhdx &
                + cfg%f_coriolis * state%v%data &
                + diff_u

        dv_dt = -(state%u%data * dvdx + state%v%data * dvdy) &
                - G * dhdy &
                - cfg%f_coriolis * state%u%data &
                + diff_v

        deallocate(dhdx, dhdy, dudx, dudy, dvdx, dvdy, &
                   dhu_dx, dhv_dy, hu, hv, diff_h, diff_u, diff_v)
    end subroutine compute_rhs

    ! 4th-order Runge-Kutta time integration
    subroutine rk4_advance(state, grid, cfg)
        type(swe_state),  intent(inout) :: state
        type(grid_t),     intent(in)    :: grid
        type(swe_config), intent(in)    :: cfg

        integer  :: nx, ny
        real(dp), allocatable :: &
            k1h(:,:), k1u(:,:), k1v(:,:), &
            k2h(:,:), k2u(:,:), k2v(:,:), &
            k3h(:,:), k3u(:,:), k3v(:,:), &
            k4h(:,:), k4u(:,:), k4v(:,:)
        type(swe_state) :: tmp
        real(dp) :: dt

        nx = grid%nx;  ny = grid%ny;  dt = cfg%dt
        allocate(k1h(nx,ny), k1u(nx,ny), k1v(nx,ny), &
                 k2h(nx,ny), k2u(nx,ny), k2v(nx,ny), &
                 k3h(nx,ny), k3u(nx,ny), k3v(nx,ny), &
                 k4h(nx,ny), k4u(nx,ny), k4v(nx,ny))

        call tmp%h%init(nx, ny, "h_tmp")
        call tmp%u%init(nx, ny, "u_tmp")
        call tmp%v%init(nx, ny, "v_tmp")

        ! k1
        call compute_rhs(state, grid, cfg, k1h, k1u, k1v)

        ! k2
        tmp%h%data = state%h%data + 0.5_dp * dt * k1h
        tmp%u%data = state%u%data + 0.5_dp * dt * k1u
        tmp%v%data = state%v%data + 0.5_dp * dt * k1v
        tmp%time   = state%time + 0.5_dp * dt
        call compute_rhs(tmp, grid, cfg, k2h, k2u, k2v)

        ! k3
        tmp%h%data = state%h%data + 0.5_dp * dt * k2h
        tmp%u%data = state%u%data + 0.5_dp * dt * k2u
        tmp%v%data = state%v%data + 0.5_dp * dt * k2v
        call compute_rhs(tmp, grid, cfg, k3h, k3u, k3v)

        ! k4
        tmp%h%data = state%h%data + dt * k3h
        tmp%u%data = state%u%data + dt * k3u
        tmp%v%data = state%v%data + dt * k3v
        tmp%time   = state%time + dt
        call compute_rhs(tmp, grid, cfg, k4h, k4u, k4v)

        ! Update
        state%h%data = state%h%data + (dt/6.0_dp) * (k1h + 2.0_dp*k2h + 2.0_dp*k3h + k4h)
        state%u%data = state%u%data + (dt/6.0_dp) * (k1u + 2.0_dp*k2u + 2.0_dp*k3u + k4u)
        state%v%data = state%v%data + (dt/6.0_dp) * (k1v + 2.0_dp*k2v + 2.0_dp*k3v + k4v)
        state%time   = state%time + dt

        deallocate(k1h, k1u, k1v, k2h, k2u, k2v, k3h, k3u, k3v, k4h, k4u, k4v)
    end subroutine rk4_advance

end module swe_numerics


module swe_io
    use swe_types
    implicit none

contains

    subroutine write_state_csv(state, grid, step)
        type(swe_state), intent(in) :: state
        type(grid_t),    intent(in) :: grid
        integer,         intent(in) :: step
        character(len=64) :: fname
        integer :: unit, i, j

        write(fname, "(a, i6.6, a)") "swe_output_", step, ".csv"
        open(newunit=unit, file=trim(fname), status="replace", action="write")
        write(unit, "(a)") "x,y,h,u,v,speed"
        do j = 1, grid%ny
            do i = 1, grid%nx
                write(unit, "(6es14.6)") &
                    grid%x(i), grid%y(j), &
                    state%h%data(i,j), &
                    state%u%data(i,j), &
                    state%v%data(i,j), &
                    sqrt(state%u%data(i,j)**2 + state%v%data(i,j)**2)
            end do
        end do
        close(unit)
        print "(a, a)", "  Written: ", trim(fname)
    end subroutine write_state_csv

    subroutine write_restart(state, grid, filename)
        type(swe_state), intent(in) :: state
        type(grid_t),    intent(in) :: grid
        character(len=*), intent(in) :: filename
        integer :: unit
        open(newunit=unit, file=trim(filename), form="unformatted", &
             status="replace", action="write")
        write(unit) state%time
        write(unit) grid%nx, grid%ny
        write(unit) state%h%data
        write(unit) state%u%data
        write(unit) state%v%data
        close(unit)
        print "(a, a)", "  Restart written: ", trim(filename)
    end subroutine write_restart

    subroutine read_restart(state, grid, filename, ok)
        type(swe_state), intent(inout) :: state
        type(grid_t),    intent(in)    :: grid
        character(len=*), intent(in)   :: filename
        logical, intent(out) :: ok
        integer :: unit, nx, ny, ios
        open(newunit=unit, file=trim(filename), form="unformatted", &
             status="old", action="read", iostat=ios)
        if (ios /= 0) then
            ok = .false.
            return
        end if
        read(unit) state%time
        read(unit) nx, ny
        if (nx /= grid%nx .or. ny /= grid%ny) then
            print *, "ERROR: restart grid mismatch"
            ok = .false.
            close(unit)
            return
        end if
        read(unit) state%h%data
        read(unit) state%u%data
        read(unit) state%v%data
        close(unit)
        ok = .true.
        print "(a, a, a, f8.1)", "  Restart loaded: ", trim(filename), "  t=", state%time
    end subroutine read_restart

    subroutine print_diagnostics(state, grid, step)
        type(swe_state), intent(in) :: state
        type(grid_t),    intent(in) :: grid
        integer,         intent(in) :: step
        real(dp) :: ke, pe, total_mass, max_speed
        integer  :: nx, ny

        nx = grid%nx;  ny = grid%ny

        ! Kinetic energy
        ke = 0.5_dp * sum(state%h%data * (state%u%data**2 + state%v%data**2)) &
             * grid%dx * grid%dy

        ! Potential energy
        pe = 0.5_dp * 9.81_dp * sum(state%h%data**2) * grid%dx * grid%dy

        ! Total mass (should be conserved)
        total_mass = sum(state%h%data) * grid%dx * grid%dy

        ! Max wind speed
        max_speed = maxval(sqrt(state%u%data**2 + state%v%data**2))

        print "(i6, f10.1, 4es12.4)", step, state%time, ke, pe, total_mass, max_speed
    end subroutine print_diagnostics

end module swe_io


program weather_model
    use swe_types
    use swe_numerics
    use swe_io
    implicit none

    type(swe_state)  :: state
    type(grid_t)     :: grid
    type(swe_config) :: cfg
    integer :: step, i, j
    real(dp) :: x, y, r2, amp
    logical :: restart_ok
    character(len=256) :: restart_file

    ! ── Configuration (would normally come from namelist file) ────────────────
    cfg%nx           = 64
    cfg%ny           = 64
    cfg%Lx           = 1.0e6_dp    ! 1000 km domain
    cfg%Ly           = 1.0e6_dp
    cfg%dt           = 60.0_dp     ! 60 second timestep
    cfg%nt           = 1440        ! 24 hours
    cfg%f_coriolis   = 1.0e-4_dp   ! mid-latitude Coriolis
    cfg%h0           = 1000.0_dp   ! 1 km mean depth
    cfg%output_freq  = 120         ! output every 2 hours
    cfg%write_restart = .true.
    cfg%output_prefix = "swe"

    ! ── Initialize grid ───────────────────────────────────────────────────────
    call grid%init(cfg%nx, cfg%ny, cfg%Lx, cfg%Ly)

    ! ── Initialize state ──────────────────────────────────────────────────────
    call state%h%init(cfg%nx, cfg%ny, "h")
    call state%u%init(cfg%nx, cfg%ny, "u")
    call state%v%init(cfg%nx, cfg%ny, "v")
    state%time = 0.0_dp

    ! Check for restart
    restart_file = "swe_restart.bin"
    call read_restart(state, grid, restart_file, restart_ok)

    if (.not. restart_ok) then
        ! Initial condition: geostrophically balanced vortex
        ! h = h0 + amp * exp(-r^2 / R^2)
        ! u, v from geostrophic balance: f*u = -g*dh/dy, f*v = g*dh/dx
        amp = 50.0_dp          ! 50 m height perturbation
        do j = 1, cfg%ny
            do i = 1, cfg%nx
                x  = grid%x(i) - cfg%Lx * 0.5_dp
                y  = grid%y(j) - cfg%Ly * 0.5_dp
                r2 = (x**2 + y**2) / (2.0e5_dp)**2   ! R = 200 km

                state%h%data(i,j) = cfg%h0 + amp * exp(-r2)

                ! Geostrophic wind
                state%u%data(i,j) = -(G * amp / cfg%f_coriolis) * &
                    (-2.0_dp * y / (2.0e5_dp)**2) * exp(-r2)
                state%v%data(i,j) =  (G * amp / cfg%f_coriolis) * &
                    (-2.0_dp * x / (2.0e5_dp)**2) * exp(-r2)
            end do
        end do
        print *, "Initialized geostrophic vortex"
    end if

    ! ── Print setup ───────────────────────────────────────────────────────────
    print *, "============================================"
    print *, " Shallow Water Equations — Weather Model"
    print *, "============================================"
    call grid%print()
    print "(a, f8.1, a)", "  dt = ", cfg%dt, " s"
    print "(a, i0, a, f8.1, a)", "  Running ", cfg%nt, " steps (", &
        cfg%nt * cfg%dt / 3600.0_dp, " hours)"
    print *, ""
    print "(a6, a10, 4a12)", "Step", "Time(s)", "KE", "PE", "Mass", "MaxSpeed"
    print "(a6, a10, 4a12)", "----", "-------", "--", "--", "----", "--------"

    ! ── Time integration ──────────────────────────────────────────────────────
    do step = 1, cfg%nt
        call rk4_advance(state, grid, cfg)

        if (mod(step, cfg%output_freq) == 0) then
            call print_diagnostics(state, grid, step)
            call write_state_csv(state, grid, step)
        end if
    end do

    ! ── Final output ──────────────────────────────────────────────────────────
    print *, ""
    print *, "=== Final State ==="
    print "(a, f8.4)", "  h: min=", state%h%min_val()
    print "(a, f8.4)", "  h: max=", state%h%max_val()
    print "(a, f8.4)", "  h: rms=", state%h%rms()
    print "(a, f8.4)", "  u: rms=", state%u%rms()
    print "(a, f8.4)", "  v: rms=", state%v%rms()

    if (cfg%write_restart) then
        call write_restart(state, grid, restart_file)
    end if

    print *, ""
    print *, "Simulation complete."

end program weather_model
