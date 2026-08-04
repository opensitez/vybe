! vybe-test: fortran/character_arrays/char_array_len
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=15) :: arr(3)
    arr(1) = 'short'
    print *, len(arr(1))
    print *, len_trim(arr(1))
end program test
