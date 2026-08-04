! vybe-test: fortran/character_arrays/char_array_decl
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=10) :: names(5)
    names(1) = 'Alice'
    print *, trim(names(1))
end program test
