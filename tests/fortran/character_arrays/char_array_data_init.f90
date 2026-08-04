! vybe-test: fortran/character_arrays/char_array_data_init
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=5) :: fruits(3)
    data fruits /'apple', 'mango', 'grape'/
    print *, trim(fruits(2))
end program test
