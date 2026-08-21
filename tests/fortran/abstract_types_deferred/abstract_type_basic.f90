! vybe-test: fortran/abstract_types_deferred/abstract_type_basic
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    type, abstract :: Shape
        real :: color(3)
    contains
        procedure(area_iface), deferred :: area
    end type Shape
    print *, "ok"
end program test

abstract interface
    function area_iface(self) result(a)
        import Shape
        class(Shape), intent(in) :: self
        real :: a
    end function area_iface
end interface
