! vybe-test: fortran/character_arrays/char_array_adjustl
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=10) :: arr(2)
    arr(1) = '   hi'
    arr(2) = ' world'
    print *, trim(adjustl(arr(1)))
    print *, trim(adjustl(arr(2)))
end program test
