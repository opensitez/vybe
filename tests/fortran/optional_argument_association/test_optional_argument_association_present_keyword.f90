! vybe-test: fortran/optional_argument_association/test_optional_argument_association_present_keyword
! origin: languages/fortran/tests/fortran/test_optional_argument_association.rs

program test_optional_argument_association
    call show(5)
    call show(5, 2)

contains
    subroutine show(a, b)
        integer, intent(in) :: a
        integer, optional, intent(in) :: b
        if (present(b)) then
            if ((a + b) /= 5) then
    print *, "FAIL: want [5] got [", a + b, "]"
    stop 1
end if
        else
            if ((a) /= 7) then
    print *, "FAIL: want [7] got [", a, "]"
    stop 1
end if
        end if
    end subroutine
end program test_optional_argument_association
