! vybe-test: fortran/fortran2003/polymorphic_arg_in
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    type :: Animal
        character(len=20) :: name = 'unknown'
    end type Animal
    type, extends(Animal) :: Dog
    end type Dog
    type(Dog) :: d
    d%name = 'Rex'
    call show_name(d)
contains
    subroutine show_name(a)
        class(Animal), intent(in) :: a
        print *, trim(a%name)
    end subroutine show_name
end program test
