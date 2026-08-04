! vybe-test: fortran/legacy_data_extended/data_implied_do_in_subroutine
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    call init()
contains
    subroutine init()
        integer :: buf(3)
        data (buf(i), i = 1, 3) /4, 5, 6/
        print *, buf(2)
    end subroutine init
end program t
