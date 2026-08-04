! vybe-test: fortran/variable_declarations_extended/double_precision_assign_after_decl
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
double precision :: d
d = 4.0d0
if ((d) /= 4) then
    print *, "FAIL: want [4] got [", d, "]"
    stop 1
end if
end program t
