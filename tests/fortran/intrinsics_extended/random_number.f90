! vybe-test: fortran/intrinsics_extended/random_number
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs

program test
    real :: r
    call random_number(r)
    print *, r >= 0.0 .and. r < 1.0
end program test
