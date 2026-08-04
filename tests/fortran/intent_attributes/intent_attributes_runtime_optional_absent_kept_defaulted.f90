! vybe-test: fortran/intent_attributes/intent_attributes_runtime_optional_absent_kept_defaulted
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs

program test_intent_attributes
integer :: x = 1
call log_value(x)
call log_value(2)
contains
subroutine log_value(x, scale)
integer, intent(in) :: x
integer, optional, intent(in) :: scale
if (present(scale)) then
if ((x * scale) /= 1) then
    print *, "FAIL: want [1] got [", x * scale, "]"
    stop 1
end if
else
if ((x) /= 4) then
    print *, "FAIL: want [4] got [", x, "]"
    stop 1
end if
end if
end subroutine log_value
end program test_intent_attributes
