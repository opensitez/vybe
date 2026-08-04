! vybe-test: fortran/variable_declarations_extended/integer_kind_2_init
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer(kind=2) :: s = 300
if ((s) /= 300) then
    print *, "FAIL: want [300] got [", s, "]"
    stop 1
end if
end program t
