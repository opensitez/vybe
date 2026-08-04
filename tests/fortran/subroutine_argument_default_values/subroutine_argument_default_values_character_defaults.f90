! vybe-test: fortran/subroutine_argument_default_values/subroutine_argument_default_values_character_defaults
! origin: languages/fortran/tests/fortran/test_subroutine_argument_default_values.rs

program subroutine_argument_default_values_character_defaults
    character(len=16) :: a
    character(len=16) :: b
    call build_label(a)
    call build_label(b, 'x')
    if (trim(trim(a)) /= "root") then
    print *, "FAIL: want [root] got [", trim(a), "]"
    stop 1
end if
    if (trim(trim(b)) /= "rootx") then
    print *, "FAIL: want [rootx] got [", trim(b), "]"
    stop 1
end if
contains
    subroutine build_label(out, suffix)
        character(len=*), intent(out) :: out
        character(len=*), intent(in), optional :: suffix
        if (present(suffix)) then
            out = trim('root') // trim(suffix)
        else
            out = 'root'
        end if
    end subroutine build_label
end program subroutine_argument_default_values_character_defaults
