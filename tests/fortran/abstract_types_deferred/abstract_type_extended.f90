! vybe-test: fortran/abstract_types_deferred/abstract_type_extended
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

module shapes
    implicit none
    type, abstract :: Shape
    contains
        procedure(compute_area), deferred :: area
    end type Shape

    abstract interface
        function compute_area(self) result(a)
            import Shape
            class(Shape), intent(in) :: self
            real :: a
        end function compute_area
    end interface

    type, extends(Shape) :: Circle
        real :: radius
    contains
        procedure :: area => circle_area
    end type Circle

contains
    function circle_area(self) result(a)
        class(Circle), intent(in) :: self
        real :: a
        a = 3.14159 * self%radius ** 2
    end function circle_area
end module shapes

program test
    use shapes
    type(Circle) :: c
    c%radius = 5.0
    print *, c%area()
end program test
