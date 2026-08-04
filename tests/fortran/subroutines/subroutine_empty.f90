! vybe-test: fortran/subroutines/subroutine_empty
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program t
call greet()
contains
subroutine greet()
if (trim("hi") /= "hi") then
    print *, "FAIL: want [hi] got [", "hi", "]"
    stop 1
end if
end subroutine greet
end program t
