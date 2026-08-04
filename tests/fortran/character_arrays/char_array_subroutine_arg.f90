! vybe-test: fortran/character_arrays/char_array_subroutine_arg
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=10) :: names(4)
    names = ['Alice     ', 'Bob       ', 'Charlie   ', 'Diana     ']
    call print_names(names)
contains
    subroutine print_names(arr)
        character(len=*), intent(in) :: arr(:)
        integer :: i
        do i = 1, size(arr)
            print *, trim(arr(i))
        end do
    end subroutine print_names
end program test
