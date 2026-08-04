! vybe-test: fortran/character_arrays/char_array_fixed_len
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=1) :: letters(5) = ['a', 'b', 'c', 'd', 'e']
    print *, letters(3)
end program test
