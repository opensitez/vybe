! vybe-test: fortran/sql_cics/sql_prepare_execute_24
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
character(len=40) :: stmt
integer :: rc
stmt = 'INSERT INTO T(ID) VALUES(:id)'
EXEC SQL PREPARE S2 FROM :stmt END-EXEC
EXEC SQL EXECUTE S2 USING SQLCA END-EXEC
EXEC SQL GET DIAGNOSTICS :rc = RETURN_STATUS END-EXEC
end program p
