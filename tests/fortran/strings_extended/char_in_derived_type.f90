! vybe-test: fortran/strings_extended/char_in_derived_type
! origin: languages/fortran/tests/fortran/test_strings_extended.rs

program test
    type :: Person
        character(len=30) :: name
        integer :: age
    end type Person
    type(Person) :: p
    p%name = 'Alice'
    p%age = 30
    print *, trim(p%name)
end program test
