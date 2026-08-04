! vybe-test: fortran/legacy/assigned_goto
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: label
    assign 10 to label
    goto label
    print *, 'skipped'
10  continue
    print *, 'ok'
end program test
