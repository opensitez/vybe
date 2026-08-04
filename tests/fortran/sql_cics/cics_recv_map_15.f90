! vybe-test: fortran/sql_cics/cics_recv_map_15
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC CICS RECEIVE MAP('M1') MAPSET('S1') END-EXEC
end program p
