! vybe-test: fortran/sql_cics/cics_send_map_14
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC CICS SEND MAP('M1') MAPSET('S1') END-EXEC
end program p
