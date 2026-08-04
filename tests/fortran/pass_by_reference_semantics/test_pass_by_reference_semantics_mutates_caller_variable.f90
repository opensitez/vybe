! vybe-test: fortran/pass_by_reference_semantics/test_pass_by_reference_semantics_mutates_caller_variable
! origin: languages/fortran/tests/fortran/test_pass_by_reference_semantics.rs

program test_pass_by_reference_semantics
    integer :: value
    value = 1
    call bump(value)
    if ((value) /= 6) then
    print *, "FAIL: want [6] got [", value, "]"
    stop 1
end if

contains
    subroutine bump(x)
        integer, intent(inout) :: x
        x = x + 5
    end subroutine
end program test_pass_by_reference_semantics
