! vybe-test: fortran/legacy/entry_basic
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    call init_and_run()
contains
    subroutine init_and_run()
        print *, 'init'
        return
    entry run_only()
        print *, 'run'
    end subroutine init_and_run
end program test
