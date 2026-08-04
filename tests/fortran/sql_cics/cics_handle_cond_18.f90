! vybe-test: fortran/sql_cics/cics_handle_cond_18
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC CICS HANDLE CONDITION ERROR(10) END-EXEC
10 continue
end program p
