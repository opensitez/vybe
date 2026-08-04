! vybe-test: fortran/character_arrays/char_array_trim
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=20) :: arr(3)
    arr(1) = 'hello     '
    arr(2) = 'world  '
    arr(3) = 'foo'
    print *, len_trim(arr(1))
    print *, len_trim(arr(2))
end program test
