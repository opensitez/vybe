! vybe-test: fortran/generic_interfaces/compile_generic_read_write_bindings
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    type :: Pair
        integer :: a, b
    contains
        procedure :: read_pair
        procedure :: write_pair
        generic :: read(formatted) => read_pair
        generic :: write(formatted) => write_pair
    end type Pair
    type(Pair) :: p
    p%a = 1
    p%b = 2
    print *, p%a + p%b
contains
    subroutine read_pair(self, unit, iostat)
        class(Pair), intent(out) :: self
        integer, intent(in) :: unit
        integer, intent(out), optional :: iostat
        read(unit, *, iostat=iostat) self%a, self%b
    end subroutine read_pair
    subroutine write_pair(self, unit, iostat)
        class(Pair), intent(in) :: self
        integer, intent(in) :: unit
        integer, intent(out), optional :: iostat
        write(unit, *, iostat=iostat) self%a, self%b
    end subroutine write_pair
end program t
