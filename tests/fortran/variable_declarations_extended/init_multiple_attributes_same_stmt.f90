! vybe-test: fortran/variable_declarations_extended/init_multiple_attributes_same_stmt
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer :: a = 5, b = 7
if ((a + b) /= 12) then
    print *, "FAIL: want [12] got [", a + b, "]"
    stop 1
end if
end program t
