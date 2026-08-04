! vybe-test: fortran/coarrays/stopped_images_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    integer, allocatable :: si(:)
    si = stopped_images()
    print *, size(si)
end program test
