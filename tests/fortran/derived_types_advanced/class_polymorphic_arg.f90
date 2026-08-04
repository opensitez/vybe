! vybe-test: fortran/derived_types_advanced/class_polymorphic_arg
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Animal
        character(len=20) :: name
    end type Animal
    type, extends(Animal) :: Dog
        character(len=10) :: breed
    end type Dog
    type(Dog) :: d
    d%name = 'Rex'
    d%breed = 'Labrador'
    call describe(d)
contains
    subroutine describe(a)
        class(Animal), intent(in) :: a
        print *, trim(a%name)
    end subroutine describe
end program test
