! vybe-test: fortran/coarrays/image_status_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    integer :: status
    status = image_status(1)
    print *, status
end program test
