! vybe-test: fortran/legacy_data_extended/entry_multiple_in_function
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    call host(3)
contains
    subroutine host(n)
        integer, intent(in) :: n
        print *, n
        return
    entry host_alt(n)
        integer :: n
        print *, n + 10
    end subroutine host
end program t
