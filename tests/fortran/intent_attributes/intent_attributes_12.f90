! vybe-test: fortran/intent_attributes/intent_attributes_12
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
subroutine s(x)
integer, intent(in) :: x
integer :: y
y = x + 1
end subroutine s
