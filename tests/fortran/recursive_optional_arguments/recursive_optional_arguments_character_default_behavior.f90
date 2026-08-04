! vybe-test: fortran/recursive_optional_arguments/recursive_optional_arguments_character_default_behavior
! origin: languages/fortran/tests/fortran/test_recursive_optional_arguments.rs

program recursive_optional_arguments_character_default_behavior
    if (trim(letters('b')) /= "b") then
    print *, "FAIL: want [b] got [", letters('b'), "]"
    stop 1
end if
    if (trim(letters('b', 2)) /= "b") then
    print *, "FAIL: want [b] got [", letters('b', 2), "]"
    stop 1
end if
contains
    recursive character(len=16) function letters(ch, repeat_count) result(out)
        character(len=*), intent(in) :: ch
        integer, optional, intent(in) :: repeat_count
        if (present(repeat_count) .and. repeat_count > 0) then
            if (len_trim(ch) >= 2) then
                out = ch // '_'
            else
                out = trim(letters(ch, repeat_count - 1))
            end if
        else
            out = ch
        end if
    end function letters
end program recursive_optional_arguments_character_default_behavior
