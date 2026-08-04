! vybe-test: fortran/pointer_target_status/test_pointer_target_status_checks_association
! origin: languages/fortran/tests/fortran/test_pointer_target_status.rs

program test_pointer_target_status
    integer, target :: storage
    integer, pointer :: p
    p => storage
    if ((associated(p)) .neqv. .true.) then
    print *, "FAIL: want [True] got [", associated(p), "]"
    stop 1
end if
    nullify(p)
    if ((associated(p)) .neqv. .false.) then
    print *, "FAIL: want [False] got [", associated(p), "]"
    stop 1
end if
end program test_pointer_target_status
