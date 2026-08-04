! vybe-test: fortran/sql_cics/sql_connect_25
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
character(len=20) :: dsn
integer :: ret
EXEC SQL CONNECT TO :dsn USER 'SYS' USING 'pass' END-EXEC
EXEC SQL GET DIAGNOSTICS :ret = RETURN_STATUS END-EXEC
end program p
