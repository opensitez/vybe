! vybe-test: fortran/derived_types_advanced/tbp_call_with_swe_io_module
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

module swe_io
    use swe_types
    implicit none
contains
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
    end subroutine read_restart
end module swe_io

program test
    use swe_types
    use swe_numerics
    use swe_io
    implicit none
    type(swe_state) :: restart_state, state
    type(grid_t)    :: restart_grid, grid
    logical :: restart_ok

    call restart_grid%init(2, 2, 2.0e6_dp, 2.0e6_dp)
    call restart_state%h%init(2, 2, "h")
    call restart_state%u%init(2, 2, "u")
    call restart_state%v%init(2, 2, "v")
    restart_state%time = 12.0_dp
    restart_state%h%data = 1.0_dp
    restart_state%u%data = 2.0_dp
    restart_state%v%data = 3.0_dp
    call write_restart(restart_state, restart_grid, "swe_restart_tbp_io_negative.bin")

    call grid%init(4, 4, 1.0e6_dp, 1.0e6_dp)
    call state%h%init(4, 4, "h")
    call state%u%init(4, 4, "u")
    call state%v%init(4, 4, "v")
    state%time = 0.0_dp

    call read_restart(state, grid, "swe_restart_tbp_io_negative.bin", restart_ok)

    if (.not. restart_ok) then
        print *, "no restart"
    end if

    print *, "before print"
    call grid%print()
    print *, "after print"
end program test
