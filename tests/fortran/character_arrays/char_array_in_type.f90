! vybe-test: fortran/character_arrays/char_array_in_type
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    type :: Person
        character(len=20) :: name
        integer :: tags(3)
        character(len=10) :: labels(3)
    end type Person
    type(Person) :: p
    p%name = 'Alice'
    p%labels(1) = 'engineer'
    p%labels(2) = 'pilot'
    p%labels(3) = 'runner'
    print *, trim(p%name)
    print *, trim(p%labels(2))
end program test
