! vybe-test: fortran/coarrays/critical_coarray_update
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: counter[*]
    counter = 0
    sync all
    critical
        counter[1] = counter[1] + 1
    end critical
    sync all
    if (this_image() == 1) print *, counter
end program test
