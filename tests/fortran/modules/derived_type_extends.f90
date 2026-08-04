! vybe-test: fortran/modules/derived_type_extends
! origin: languages/fortran/tests/fortran/test_modules.rs
program t
type :: Shape
real :: area
end type Shape
type, extends(Shape) :: Circle
real :: radius
end type Circle
type(Circle) :: c
c%radius = 5.0
print *, c%radius
end program t
