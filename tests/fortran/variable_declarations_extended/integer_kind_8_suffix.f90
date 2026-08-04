! vybe-test: fortran/variable_declarations_extended/integer_kind_8_suffix
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer(kind=8) :: big = 42_8
if ((big) /= 42) then
    print *, "FAIL: want [42] got [", big, "]"
    stop 1
end if
end program t
