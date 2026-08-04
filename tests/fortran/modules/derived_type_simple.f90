! vybe-test: fortran/modules/derived_type_simple
! origin: languages/fortran/tests/fortran/test_modules.rs
program t
type :: Point
real :: x
real :: y
end type Point
type(Point) :: p
p%x = 3.0
p%y = 4.0
print *, p%x
end program t
