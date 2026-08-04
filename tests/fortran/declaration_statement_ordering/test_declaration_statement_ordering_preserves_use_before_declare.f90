! vybe-test: fortran/declaration_statement_ordering/test_declaration_statement_ordering_preserves_use_before_declare
! origin: languages/fortran/tests/fortran/test_declaration_statement_ordering.rs

program test_declaration_statement_ordering
    integer :: value
    real :: ratio
    integer, parameter :: offset = 1

    value = 10
    ratio = real(value + offset) / 2.0

    if ((value) /= 10) then
    print *, "FAIL: want [10] got [", value, "]"
    stop 1
end if
    if ((offset) /= 1) then
    print *, "FAIL: want [1] got [", offset, "]"
    stop 1
end if
    if ((nint(ratio)) /= 5) then
    print *, "FAIL: want [5] got [", nint(ratio), "]"
    stop 1
end if
end program test_declaration_statement_ordering
