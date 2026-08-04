! vybe-test: fortran/derived_types_advanced/weather_model_full_reproduction
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs
integer :: vybe_check_i = 0
character(len=18) :: vybe_check_w(4) = [ "Initialized vortex", "Header", "Grid: 4 x 4", "Done" ]

module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)
    real(dp), parameter :: G = 9.81_dp

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: dx, dy, Lx, Ly
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
        procedure :: init => field_init
    end type field2d

    type :: swe_state
        type(field2d) :: h
        type(field2d) :: u
        type(field2d) :: v
        real(dp) :: time
    end type swe_state

    type :: swe_config
        integer  :: nx, ny
        real(dp) :: Lx, Ly, dt, f_coriolis, h0
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
        print "(a, i0, a, i0)", "Grid: ", self%nx, " x ", self%ny
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
end module swe_types

module swe_numerics
    use swe_types
    implicit none
contains
    pure function wrap(i, n) result(j)
        integer, intent(in) :: i, n
        integer :: j
        j = mod(i - 1 + n, n) + 1
    end function wrap
end module swe_numerics

program weather_model
    use swe_types
    use swe_numerics
    implicit none

    type(swe_state)  :: state
    type(grid_t)     :: grid
    type(swe_config) :: cfg
    logical :: restart_ok

    cfg%nx = 4;  cfg%ny = 4
    cfg%Lx = 1.0e6_dp;  cfg%Ly = 1.0e6_dp
    cfg%dt = 60.0_dp
    cfg%f_coriolis = 1.0e-4_dp
    cfg%h0 = 1000.0_dp

    call grid%init(cfg%nx, cfg%ny, cfg%Lx, cfg%Ly)
    call state%h%init(cfg%nx, cfg%ny, "h")
    call state%u%init(cfg%nx, cfg%ny, "u")
    call state%v%init(cfg%nx, cfg%ny, "v")
    state%time = 0.0_dp

    restart_ok = .false.

    if (.not. restart_ok) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 4) then
            print *, "FAIL: more than 4 line(s)"
            stop 1
        end if
        if (trim("Initialized vortex") /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", "Initialized vortex", "]"
            stop 1
        end if
    end if

        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 4) then
        print *, "FAIL: more than 4 line(s)"
        stop 1
    end if
    if (trim("Header") /= trim(vybe_check_w(vybe_check_i))) then
        print *, "FAIL at ", vybe_check_i, " got [", "Header", "]"
        stop 1
    end if
    call grid%print()
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 4) then
        print *, "FAIL: more than 4 line(s)"
        stop 1
    end if
    if (trim("Done") /= trim(vybe_check_w(vybe_check_i))) then
        print *, "FAIL at ", vybe_check_i, " got [", "Done", "]"
        stop 1
    end if
if (vybe_check_i /= 4) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 4"
    stop 1
end if
end program weather_model
