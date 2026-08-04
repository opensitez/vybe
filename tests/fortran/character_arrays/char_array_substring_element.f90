! vybe-test: fortran/character_arrays/char_array_substring_element
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=10) :: arr(3)
    arr(1) = 'Hello World'
    print *, arr(1)(1:5)
end program test
