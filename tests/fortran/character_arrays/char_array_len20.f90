! vybe-test: fortran/character_arrays/char_array_len20
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=20) :: words(4)
    words(1) = 'Fortran'
    words(2) = 'is'
    words(3) = 'still'
    words(4) = 'relevant'
    print *, trim(words(1)), ' ', trim(words(2))
end program test
