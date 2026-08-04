! vybe-test: fortran/character_arrays/char_array_index
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=20) :: arr(3)
    arr(1) = 'hello world'
    arr(2) = 'fortran 90'
    arr(3) = 'no match here'
    print *, index(arr(1), 'world')
    print *, index(arr(3), 'xyz')
end program test
