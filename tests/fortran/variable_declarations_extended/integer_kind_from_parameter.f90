! vybe-test: fortran/variable_declarations_extended/integer_kind_from_parameter
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, parameter :: ik = 4
integer(kind=ik) :: n = 99
if ((n) /= 99) then
    print *, "FAIL: want [99] got [", n, "]"
    stop 1
end if
end program t
