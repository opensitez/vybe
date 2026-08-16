! vybe-test: fortran/derived_type_oop_extended/compile_tbp_polymorphic_child_override_dispatch
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs
module m
    type :: Animal
    contains
        procedure :: legs
    end type Animal
    type, extends(Animal) :: Spider
    contains
        procedure :: legs => spider_legs
    end type Spider
contains
    integer function legs(self) result(n)
        class(Animal), intent(in) :: self
        n = 4
    end function legs
    integer function spider_legs(self) result(n)
        class(Spider), intent(in) :: self
        n = 8
    end function spider_legs
end module m
program driver
use m
    class(Animal), allocatable :: a
    allocate(Spider :: a)
    print *, a%legs()
end program driver
