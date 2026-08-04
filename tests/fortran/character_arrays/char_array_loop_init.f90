! vybe-test: fortran/character_arrays/char_array_loop_init
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=5) :: a(3)
    integer :: i
    a(1) = 'one  '
    a(2) = 'two  '
    a(3) = 'three'
    do i = 1, 3
        print *, trim(a(i))
    end do
end program test
