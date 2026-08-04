! vybe-test: fortran/character_arrays/char_array_allocatable
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=20), allocatable :: arr(:)
    allocate(arr(5))
    arr(1) = 'first'
    arr(5) = 'last'
    print *, trim(arr(1))
    print *, trim(arr(5))
    deallocate(arr)
end program test
