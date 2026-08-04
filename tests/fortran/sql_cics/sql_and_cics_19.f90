! vybe-test: fortran/sql_cics/sql_and_cics_19
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
integer :: id
EXEC SQL SELECT 1 INTO :id END-EXEC
EXEC CICS RETURN END-EXEC
end program p
