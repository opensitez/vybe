! vybe-test: fortran/derived_types_advanced/weather_model_grid_print_tbp
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)
    real(dp), parameter :: G = 9.81_dp

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: dx, dy
        real(dp) :: Lx, Ly
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t

contains
    subroutine grid_init(self, nx, ny, Lx, Ly)
        class(grid_t), intent(inout) :: self
        integer, intent(in) :: nx, ny
        real(dp), intent(in) :: Lx, Ly
        self%nx = nx
        self%ny = ny
        self%Lx = Lx
        self%Ly = Ly
        self%dx = Lx / nx
        self%dy = Ly / ny
    end subroutine grid_init

    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        if (trim(self%nx) /= "before print") then
    print *, "FAIL: want [before print] got [", self%nx, "]"
    stop 1
end if
        if ((self%ny) /= 10) then
    print *, "FAIL: want [10] got [", self%ny, "]"
    stop 1
end if
    end subroutine grid_print
end module swe_types

program test_weather
    use swe_types
    type(grid_t) :: grid
    call grid%init(10, 10, 1000.0d0, 1000.0d0)
    if (("before print") /= 10) then
    print *, "FAIL: want [10] got [", "before print", "]"
    stop 1
end if
    call grid%print()
    if (trim("after print") /= "after print") then
    print *, "FAIL: want [after print] got [", "after print", "]"
    stop 1
end if
end program test_weather
