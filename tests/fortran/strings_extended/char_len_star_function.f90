! vybe-test: fortran/strings_extended/char_len_star_function
! origin: languages/fortran/tests/fortran/test_strings_extended.rs

program test
    call show('hello')
contains
    subroutine show(msg)
        character(len=*), intent(in) :: msg
        if (trim(trim(msg)) /= "hello") then
    print *, "FAIL: want [hello] got [", trim(msg), "]"
    stop 1
end if
    end subroutine
end program test
