! vybe-test: fortran/legacy_data_extended/entry_with_dummy_arguments
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    call master(5)
contains
    subroutine master(x)
        integer, intent(in) :: x
        print *, x
        return
    entry slave(y)
        integer :: y
        print *, y + 1
    end subroutine master
end program t
