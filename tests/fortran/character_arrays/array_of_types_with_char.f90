! vybe-test: fortran/character_arrays/array_of_types_with_char
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    type :: Tag
        character(len=20) :: name
        integer :: value
    end type Tag
    type(Tag) :: tags(3)
    tags(1)%name = 'alpha'
    tags(1)%value = 1
    tags(2)%name = 'beta'
    tags(2)%value = 2
    tags(3)%name = 'gamma'
    tags(3)%value = 3
    print *, trim(tags(2)%name)
end program test
