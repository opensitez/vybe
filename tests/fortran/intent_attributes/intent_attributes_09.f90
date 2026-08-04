! vybe-test: fortran/intent_attributes/intent_attributes_09
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
subroutine s(x)
complex, intent(inout) :: x
end subroutine s
