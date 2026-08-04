! vybe-test: fortran/character_arrays/char_array_substring_assign
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=10) :: arr(2)
    arr(1) = 'XXXXXXXXXX'
    arr(1)(1:5) = 'Hello'
    print *, trim(arr(1))
end program test
