! vybe-test: fortran/derived_types_advanced/array_of_types
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Point
        real :: x, y
    end type Point
    type(Point) :: pts(3)
    integer :: i
    do i = 1, 3
        pts(i)%x = real(i)
        pts(i)%y = real(i) * 2.0
    end do
    print *, pts(2)%x
end program test
