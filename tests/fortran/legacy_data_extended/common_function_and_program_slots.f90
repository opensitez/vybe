! vybe-test: fortran/legacy_data_extended/common_function_and_program_slots
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs

program t
    integer :: n
    real :: rate
    common /cfg/ n, rate
    n = 5
    rate = 2.0
    print *, scaled()
contains
    function scaled()
        real :: scaled
        integer :: n
        real :: rate
        common /cfg/ n, rate
        scaled = n * rate
    end function scaled
end program t
