! vybe-test: fortran/variable_declarations_extended/logical_kind_from_not_expression
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
logical :: a = .true.
logical :: b = .not. a
if ((b) .neqv. .false.) then
    print *, "FAIL: want [false] got [", b, "]"
    stop 1
end if
end program t
