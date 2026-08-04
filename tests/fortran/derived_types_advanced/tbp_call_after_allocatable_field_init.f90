! vybe-test: fortran/derived_types_advanced/tbp_call_after_allocatable_field_init
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: dx, dy, Lx, Ly
        real(dp), allocatable :: x(:), y(:)
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t

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
        if ((self%nx) /= 4) then
    print *, "FAIL: want [4] got [", self%nx, "]"
    stop 1
end if
        if ((self%ny) /= 4) then
    print *, "FAIL: want [4] got [", self%ny, "]"
    stop 1
end if
    end subroutine grid_print
end module swe_types

program test
    use swe_types
    type(grid_t) :: grid
    call grid%init(4, 4, 1000.0d0, 1000.0d0)
    call grid%print()
end program test
