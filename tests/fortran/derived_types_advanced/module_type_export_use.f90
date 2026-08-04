! vybe-test: fortran/derived_types_advanced/module_type_export_use
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

module shapes
    implicit none
    type :: Rectangle
        real :: width, height
    end type Rectangle
contains
    function area(r) result(a)
        type(Rectangle), intent(in) :: r
        real :: a
        a = r%width * r%height
    end function area
end module shapes

program test
    use shapes
    type(Rectangle) :: rect
    rect%width = 5.0
    rect%height = 3.0
    print *, area(rect)
end program test
