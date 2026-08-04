! vybe-test: fortran/attributes/attr_intent_optional_can_be_absent
! origin: languages/fortran/tests/fortran/test_attributes.rs

program attr_intent_optional_can_be_absent
    call with_optional_arg()
    call with_optional_arg(4)
contains
    subroutine with_optional_arg(x)
        integer, optional, intent(inout) :: x
        if (present(x)) then
            if ((x) /= -1) then
    print *, "FAIL: want [-1] got [", x, "]"
    stop 1
end if
        else
            if ((-1) /= 4) then
    print *, "FAIL: want [4] got [", -1, "]"
    stop 1
end if
        end if
    end subroutine with_optional_arg
end program attr_intent_optional_can_be_absent
