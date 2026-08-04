! vybe-test: fortran/legacy_data_extended/save_in_function_result
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    print *, counter()
    print *, counter()
contains
    function counter()
        integer, save :: n = 0
        integer :: counter
        n = n + 1
        counter = n
    end function counter
end program t
