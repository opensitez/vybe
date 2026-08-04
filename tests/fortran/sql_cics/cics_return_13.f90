! vybe-test: fortran/sql_cics/cics_return_13
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC CICS RETURN END-EXEC
end program p
