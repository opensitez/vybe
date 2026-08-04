! vybe-test: fortran/legacy_data_extended/entry_nested_return_paths
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    call pipeline()
contains
    subroutine pipeline()
        print *, 1
        return
    entry pipeline_b()
        print *, 2
        return
    entry pipeline_c()
        print *, 3
    end subroutine pipeline
end program t
