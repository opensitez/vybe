! vybe-test: fortran/sql_cics/cics_handle_cond_multi_27
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC CICS HANDLE CONDITION LTERM(10) INVALIDP(20) ENDPGM(30) END-EXEC
10 continue
20 continue
30 continue
end program p
