! vybe-test: fortran/character_arrays/char_array_element_assign
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=10) :: arr(3)
    arr(1) = 'hello'
    arr(2) = 'world'
    arr(3) = 'foo'
    arr(2) = 'fortran'
    print *, trim(arr(2))
end program test
