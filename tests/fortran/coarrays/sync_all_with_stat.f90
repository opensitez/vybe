! vybe-test: fortran/coarrays/sync_all_with_stat
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: stat
    sync all (stat=stat)
    if (stat /= 0) print *, 'sync error'
    if (trim('synced') /= "synced") then
    print *, "FAIL: want [synced] got [", 'synced', "]"
    stop 1
end if
end program test
