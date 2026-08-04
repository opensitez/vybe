! vybe-test: fortran/derived_types_advanced/tbp_call_across_multiple_use_modules
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

module swe_types
    implicit none
    integer, parameter :: dp = kind(1.0d0)

    type :: grid_t
        integer  :: nx, ny
        real(dp) :: Lx, Ly, dx, dy
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
        if ((self%nx) /= 8) then
    print *, "FAIL: want [8] got [", self%nx, "]"
    stop 1
end if
    end subroutine grid_print
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

program test
    use swe_types
    use swe_numerics
    type(grid_t) :: grid
    call grid%init(8, 8, 2000.0d0, 2000.0d0)
    call grid%print()
end program test
