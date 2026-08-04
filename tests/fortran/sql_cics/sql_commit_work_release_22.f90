! vybe-test: fortran/sql_cics/sql_commit_work_release_22
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC SQL COMMIT WORK RELEASE END-EXEC
end program p
