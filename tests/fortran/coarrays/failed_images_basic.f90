! vybe-test: fortran/coarrays/failed_images_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    use iso_fortran_env
    integer, allocatable :: fi(:)
    fi = failed_images()
    print *, size(fi)
end program test
