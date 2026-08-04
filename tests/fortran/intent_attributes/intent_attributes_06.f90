! vybe-test: fortran/intent_attributes/intent_attributes_06
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
subroutine s(a)
real, intent(inout) :: a(:)
end subroutine s
