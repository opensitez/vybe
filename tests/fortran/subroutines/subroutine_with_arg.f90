! vybe-test: fortran/subroutines/subroutine_with_arg
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program t
call say("hello")
contains
subroutine say(msg)
character(len=*), intent(in) :: msg
if (trim(msg) /= "hello") then
    print *, "FAIL: want [hello] got [", msg, "]"
    stop 1
end if
end subroutine say
end program t
