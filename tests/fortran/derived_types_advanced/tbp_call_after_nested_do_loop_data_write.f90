! vybe-test: fortran/derived_types_advanced/tbp_call_after_nested_do_loop_data_write
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

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
        print *, self%nx
        print *, self%ny
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

program test
    use swe_types
    implicit none
    type(swe_state) :: state
    type(grid_t)    :: grid
    integer :: i, j
    real(dp) :: x, y, r2, amp
    integer, parameter :: nx = 4, ny = 4
    real(dp), parameter :: Lx = 1.0e6_dp, Ly = 1.0e6_dp
    real(dp), parameter :: f = 1.0e-4_dp, h0 = 1000.0_dp

    call grid%init(nx, ny, Lx, Ly)
    call state%h%init(nx, ny, "h")
    call state%u%init(nx, ny, "u")
    call state%v%init(nx, ny, "v")
    state%time = 0.0_dp

    ! Geostrophic vortex initial condition (the nested do-loops from weather_model)
    amp = 50.0_dp
    do j = 1, ny
        do i = 1, nx
            x  = grid%x(i) - Lx * 0.5_dp
            y  = grid%y(j) - Ly * 0.5_dp
            r2 = (x**2 + y**2) / (2.0e5_dp)**2

            state%h%data(i,j) = h0 + amp * exp(-r2)

            state%u%data(i,j) = -(G * amp / f) * &
                (-2.0_dp * y / (2.0e5_dp)**2) * exp(-r2)
            state%v%data(i,j) =  (G * amp / f) * &
                (-2.0_dp * x / (2.0e5_dp)**2) * exp(-r2)
        end do
    end do

    print *, "vortex done"
    call grid%print()
    print *, "print done"
end program test
