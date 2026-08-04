! vybe-test: fortran/legacy_data_extended/save_persists_across_calls
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    call tick()
    call tick()
    call tick()
contains
    subroutine tick()
        integer, save :: n = 0
        n = n + 1
        print *, n
    end subroutine tick
end program t
