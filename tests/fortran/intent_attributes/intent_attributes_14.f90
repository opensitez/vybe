! vybe-test: fortran/intent_attributes/intent_attributes_14
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program p
integer :: i
call s(i)
contains
subroutine s(a)
integer, intent(inout) :: a
a = a + 1
end subroutine s
end program p
