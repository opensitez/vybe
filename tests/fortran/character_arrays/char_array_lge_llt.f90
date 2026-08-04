! vybe-test: fortran/character_arrays/char_array_lge_llt
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=5) :: arr(2) = ['abc  ', 'xyz  ']
    print *, llt(arr(1), arr(2))
    print *, lge(arr(2), arr(1))
end program test
