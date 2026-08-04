! vybe-test: fortran/fortran2003/polymorphic_arg_in_runtime
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
        if (trim(trim(a%name)) /= "Rex") then
    print *, "FAIL: want [Rex] got [", trim(a%name), "]"
    stop 1
end if
    end subroutine show_name
end program test
