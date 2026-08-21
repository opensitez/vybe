! vybe-test: fortran/type_bound_procedures/compile_tbp_non_overridable_binding
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module m
    type :: Fixed
        integer :: n = 1
    contains
        procedure, non_overridable :: id
    end type Fixed
contains
    function id(self) result(v)
        class(Fixed), intent(in) :: self
        integer :: v
        v = self%n
    end function id
end module m
program driver
use m
    type(Fixed) :: f
    print *, f%id()
end program driver
