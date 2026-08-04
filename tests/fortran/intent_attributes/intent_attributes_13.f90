! vybe-test: fortran/intent_attributes/intent_attributes_13
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
subroutine s(a, b, c)
integer, intent(inout) :: a
integer, intent(in) :: b
integer, intent(out) :: c
c = a + b
a = c
end subroutine s
