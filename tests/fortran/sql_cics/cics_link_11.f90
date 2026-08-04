! vybe-test: fortran/sql_cics/cics_link_11
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC CICS LINK PROGRAM('SUB1') END-EXEC
end program p
