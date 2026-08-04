! vybe-test: fortran/coarray_sync_extended/sync_all_stat_errmsg_reports_ok
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
integer :: stat
character(len=80) :: errmsg
sync all (stat=stat, errmsg=errmsg)
if (stat == 0) print *, 'barrier ok'
end program t
