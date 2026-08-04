! vybe-test: fortran/derived_types_advanced/nested_derived_types_with_tbp
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

module swe_types
    implicit none
    type :: field2d
        integer :: nx, ny
        character(len=32) :: name
    contains
        procedure :: init    => field_init
    end type field2d
    type :: swe_state
        type(field2d) :: h
        type(field2d) :: u
    end type swe_state
contains
    subroutine field_init(self, nx, ny, name)
        class(field2d), intent(inout) :: self
        integer, intent(in) :: nx, ny
        character(len=*), intent(in) :: name
        self%nx = nx
        self%ny = ny
        self%name = name
    end subroutine field_init
end module swe_types
program test
    use swe_types
    type(swe_state) :: state
    call state%h%init(4, 4, "depth")
    call state%u%init(4, 4, "u-vel")
    if ((state%h%nx) /= 4) then
    print *, "FAIL: want [4] got [", state%h%nx, "]"
    stop 1
end if
    if (trim(state%h%name) /= "depth") then
    print *, "FAIL: want [depth] got [", state%h%name, "]"
    stop 1
end if
    if (trim(state%u%name) /= "u-vel") then
    print *, "FAIL: want [u-vel] got [", state%u%name, "]"
    stop 1
end if
end program test
