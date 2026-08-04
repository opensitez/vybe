! vybe-test: fortran/intent_attributes/intent_attributes_runtime_array_inout_section_update
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs

program test_intent_attributes
integer :: a(3) = [1,2,3]
call inc_section(a(1:3:2))
if ((a(1)) /= 2) then
    print *, "FAIL: want [2] got [", a(1), "]"
    stop 1
end if
if ((a(3)) /= 4) then
    print *, "FAIL: want [4] got [", a(3), "]"
    stop 1
end if

contains
subroutine inc_section(x)
integer, intent(inout) :: x(:)
x = x + 1
end subroutine inc_section
end program test_intent_attributes
