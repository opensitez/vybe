! vybe-test: fortran/subroutine_extended/internal_sub_calls_inner_print
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
call outer()
contains
subroutine outer()
call inner(42)
end subroutine outer
subroutine inner(v)
integer, intent(in) :: v
if ((v) /= 42) then
    print *, "FAIL: want [42] got [", v, "]"
    stop 1
end if
end subroutine inner
end program t
