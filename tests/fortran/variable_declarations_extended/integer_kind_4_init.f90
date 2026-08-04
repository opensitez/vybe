! vybe-test: fortran/variable_declarations_extended/integer_kind_4_init
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer(kind=4) :: x = 17
if ((x) /= 17) then
    print *, "FAIL: want [17] got [", x, "]"
    stop 1
end if
end program t
