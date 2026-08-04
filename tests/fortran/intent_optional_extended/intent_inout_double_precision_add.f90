! vybe-test: fortran/intent_optional_extended/intent_inout_double_precision_add
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
double precision :: x
x = 1.5d0
call add_half(x)
if ((int(x)) /= 2) then
    print *, "FAIL: want [2] got [", int(x), "]"
    stop 1
end if
contains
subroutine add_half(v)
double precision, intent(inout) :: v
v = v + 0.5d0
end subroutine add_half
end program t
