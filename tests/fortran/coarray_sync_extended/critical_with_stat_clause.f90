! vybe-test: fortran/coarray_sync_extended/critical_with_stat_clause
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    integer :: n = 0
    integer :: stat
    critical (stat=stat)
        n = n + 1
    end critical
    print *, n, stat
end program t
