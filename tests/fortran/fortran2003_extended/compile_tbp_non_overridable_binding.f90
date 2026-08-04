! vybe-test: fortran/fortran2003_extended/compile_tbp_non_overridable_binding
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    type :: Fixed
        integer :: n = 1
    contains
        procedure, non_overridable :: id
    end type Fixed
    type(Fixed) :: f
    print *, f%id()
contains
    function id(self) result(v)
        class(Fixed), intent(in) :: self
        integer :: v
        v = self%n
    end function id
end program t
