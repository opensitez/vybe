! vybe-test: fortran/subroutine_extended/internal_two_subs_print_in_order
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
call first()
call second()
contains
subroutine first()
if ((1) /= 1) then
    print *, "FAIL: want [1] got [", 1, "]"
    stop 1
end if
end subroutine first
subroutine second()
if ((2) /= 2) then
    print *, "FAIL: want [2] got [", 2, "]"
    stop 1
end if
end subroutine second
end program t
