! vybe-test: fortran/legacy_data_extended/entry_function_primary_body
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    print *, primary(4)
contains
    function primary(n)
        integer, intent(in) :: n
        integer :: primary
        primary = n * 2
    entry backup(n)
        integer :: n
        primary = n + 1
    end function primary
end program t
