! vybe-test: fortran/intent_attributes/intent_attributes_11
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
subroutine s(x, y)
integer, intent(in) :: x
integer, optional, intent(in) :: y
end subroutine s
