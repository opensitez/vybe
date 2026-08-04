! vybe-test: fortran/derived_types_advanced/tbp_with_print_method_name
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

module swe_types
    implicit none
    type :: grid_t
        integer  :: nx, ny
    contains
        procedure :: init  => grid_init
        procedure :: print => grid_print
    end type grid_t
contains
    subroutine grid_init(self, nx, ny)
        class(grid_t), intent(inout) :: self
        integer, intent(in) :: nx, ny
        self%nx = nx
        self%ny = ny
    end subroutine grid_init
    subroutine grid_print(self)
        class(grid_t), intent(in) :: self
        if ((self%nx) /= 10) then
    print *, "FAIL: want [10] got [", self%nx, "]"
    stop 1
end if
        if ((self%ny) /= 20) then
    print *, "FAIL: want [20] got [", self%ny, "]"
    stop 1
end if
    end subroutine grid_print
end module swe_types
program test_weather
    use swe_types
    type(grid_t) :: grid
    call grid%init(10, 20)
    call grid%print()
end program test_weather
