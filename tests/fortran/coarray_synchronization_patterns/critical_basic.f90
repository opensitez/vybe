! vybe-test: fortran/coarray_synchronization_patterns/critical_basic
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: shared = 0
    critical
        shared = shared + 1
    end critical
    print *, shared
end program test
