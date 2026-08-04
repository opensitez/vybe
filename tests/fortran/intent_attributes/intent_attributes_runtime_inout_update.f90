! vybe-test: fortran/intent_attributes/intent_attributes_runtime_inout_update
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs

program test_intent_attributes
integer :: x = 2
call bump(x)
if ((x) /= 5) then
    print *, "FAIL: want [5] got [", x, "]"
    stop 1
end if

contains
subroutine bump(v)
integer, intent(inout) :: v
v = v + 3
end subroutine bump
end program test_intent_attributes
