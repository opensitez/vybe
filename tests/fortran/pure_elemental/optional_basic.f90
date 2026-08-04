! vybe-test: fortran/pure_elemental/optional_basic
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    call greet('Alice')
    call greet('Bob', 'Dr.')
contains
    subroutine greet(name, title)
        character(len=*), intent(in) :: name
        character(len=*), intent(in), optional :: title
        if (present(title)) then
            print *, trim(title) // ' ' // trim(name)
        else
            print *, trim(name)
        end if
    end subroutine greet
end program test
