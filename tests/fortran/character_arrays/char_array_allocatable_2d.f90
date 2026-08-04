! vybe-test: fortran/character_arrays/char_array_allocatable_2d
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=10), allocatable :: grid(:,:)
    allocate(grid(3,3))
    grid(1,1) = 'top-left'
    grid(3,3) = 'bot-right'
    print *, trim(grid(1,1))
    deallocate(grid)
end program test
