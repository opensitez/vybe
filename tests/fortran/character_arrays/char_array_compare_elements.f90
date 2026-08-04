! vybe-test: fortran/character_arrays/char_array_compare_elements
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=5) :: arr(3) = ['apple', 'mango', 'grape']
    print *, arr(1) < arr(2)
    print *, arr(2) > arr(3)
    print *, arr(1) == arr(1)
end program test
