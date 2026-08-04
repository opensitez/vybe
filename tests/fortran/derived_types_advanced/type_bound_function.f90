! vybe-test: fortran/derived_types_advanced/type_bound_function
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Vector
        real :: x, y
    contains
        procedure :: magnitude
    end type Vector
    type(Vector) :: v
    v%x = 3.0
    v%y = 4.0
    print *, v%magnitude()
contains
    function magnitude(self) result(m)
        class(Vector), intent(in) :: self
        real :: m
        m = sqrt(self%x**2 + self%y**2)
    end function magnitude
end program test
