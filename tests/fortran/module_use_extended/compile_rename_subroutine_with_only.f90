! vybe-test: fortran/module_use_extended/compile_rename_subroutine_with_only
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs

module runner
    implicit none
contains
    subroutine execute_job()
        print *, "done"
    end subroutine execute_job
end module runner

program t
    use runner, only: go => execute_job
    call go()
end program t
