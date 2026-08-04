! vybe-test: fortran/legacy_data_extended/save_uninitialized_reads_zero
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
call read_zero()
contains
subroutine read_zero()
integer, save :: bucket
if ((bucket) /= 0) then
    print *, "FAIL: want [0] got [", bucket, "]"
    stop 1
end if
end subroutine read_zero
end program t
