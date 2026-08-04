! vybe-test: fortran/derived_types_advanced/weather_model_full_types
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)
    real(dp), parameter :: G = 9.81_dp

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: Lx, Ly
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t

    type :: field2d
        integer :: nx, ny
        character(len=32) :: name
    contains
        procedure :: init    => field_init
    end type field2d

    type :: swe_state
        type(field2d) :: h
        type(field2d) :: u
        type(field2d) :: v
        real(dp)      :: time
    end type swe_state

    type :: swe_config
        integer  :: nx, ny
        integer  :: nt
        real(dp) :: dt
    end type swe_config

contains
    subroutine grid_init(self, nx, ny, Lx, Ly)
        class(grid_t), intent(inout) :: self
        integer, intent(in) :: nx, ny
        real(dp), intent(in) :: Lx, Ly
        self%nx = nx
        self%ny = ny
        self%Lx = Lx
        self%Ly = Ly
    end subroutine grid_init

    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        if ((self%nx) /= 10) then
    print *, "FAIL: want [10] got [", self%nx, "]"
    stop 1
end if
        if ((self%ny) /= 10) then
    print *, "FAIL: want [10] got [", self%ny, "]"
    stop 1
end if
    end subroutine grid_print

    subroutine field_init(self, nx, ny, name)
        class(field2d), intent(inout) :: self
        integer, intent(in) :: nx, ny
        character(len=*), intent(in) :: name
        self%nx = nx
        self%ny = ny
        self%name = name
    end subroutine field_init
end module swe_types

program test_weather
    use swe_types
    type(grid_t)    :: grid
    type(swe_state) :: state
    type(swe_config) :: cfg

    call grid%init(10, 10, 1000.0d0, 1000.0d0)
    call state%h%init(10, 10, "depth")
    call grid%print()
    if (trim(state%h%name) /= "depth") then
    print *, "FAIL: want [depth] got [", state%h%name, "]"
    stop 1
end if
end program test_weather
