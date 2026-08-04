! vybe-test: fortran/legacy_data_extended/entry_call_alternate_name
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    call door()
contains
    subroutine door()
        print *, 'main'
        return
    entry window()
        print *, 'alt'
    end subroutine door
end program t
