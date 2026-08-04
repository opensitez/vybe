! vybe-test: fortran/legacy/data_logical
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    logical :: flag
    data flag /.true./
    print *, flag
end program test
