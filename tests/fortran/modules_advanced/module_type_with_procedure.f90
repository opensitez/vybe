! vybe-test: fortran/modules_advanced/module_type_with_procedure
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module animals
    implicit none
    type :: Dog
        character(len=20) :: name
    contains
        procedure :: speak
    end type Dog
contains
    subroutine speak(self)
        class(Dog), intent(in) :: self
        if (trim('Woof! I am ' // trim(self%name)) /= "Woof! I am Rex") then
    print *, "FAIL: want [Woof! I am Rex] got [", 'Woof! I am ' // trim(self%name), "]"
    stop 1
end if
    end subroutine speak
end module animals

program test
    use animals
    type(Dog) :: d
    d%name = 'Rex'
    call d%speak()
end program test
