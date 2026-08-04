! vybe-test: fortran/subroutines/multiple_contains
! origin: languages/fortran/tests/fortran/test_subroutines.rs
program t
contains
subroutine a()
if (trim("a") /= "a") then
    print *, "FAIL: want [a] got [", "a", "]"
    stop 1
end if
end subroutine a
subroutine b()
if (trim("b") /= "b") then
    print *, "FAIL: want [b] got [", "b", "]"
    stop 1
end if
end subroutine b
end program t
