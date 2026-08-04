! vybe-test: fortran/attributes/attr_intent_optional_value_flow
! origin: languages/fortran/tests/fortran/test_attributes.rs

program attr_intent_optional_value_flow
    call with_attributes(10, .true.)
contains
    subroutine with_attributes(x, flag)
        integer, optional, intent(in) :: x
        logical, intent(in) :: flag
        if ((present(x)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", present(x), "]"
    stop 1
end if
        if ((flag) .neqv. .true.) then
    print *, "FAIL: want [true] got [", flag, "]"
    stop 1
end if
    end subroutine with_attributes
end program attr_intent_optional_value_flow
