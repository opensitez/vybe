! vybe-test: fortran/legacy_data_extended/save_all_locals_listed
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    call stash(3)
    call recall()
contains
    subroutine stash(v)
        integer, intent(in) :: v
        integer, save :: a, b
        a = v
        b = v * 2
    end subroutine stash
    subroutine recall()
        integer, save :: a, b
        print *, a, b
    end subroutine recall
end program t
