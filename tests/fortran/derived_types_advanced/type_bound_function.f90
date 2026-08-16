! vybe-test: fortran/derived_types_advanced/type_bound_function
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs
module m
    type :: Vector
        real :: x, y
    contains
        procedure :: magnitude
    end type Vector
contains
    function magnitude(self) result(m)
        class(Vector), intent(in) :: self
        real :: m
        m = sqrt(self%x**2 + self%y**2)
    end function magnitude
end module m
program driver
use m
    type(Vector) :: v
    v%x = 3.0
    v%y = 4.0
    print *, v%magnitude()
end program driver
