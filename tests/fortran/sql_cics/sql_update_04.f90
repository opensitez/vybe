! vybe-test: fortran/sql_cics/sql_update_04
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
integer :: id
id = 1
EXEC SQL UPDATE T SET ID = :id END-EXEC
end program p
