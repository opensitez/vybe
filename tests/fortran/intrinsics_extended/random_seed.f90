! vybe-test: fortran/intrinsics_extended/random_seed
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs

program test
    integer :: seed(1) = [42]
    call random_seed(put=seed)
    print *, "ok"
end program test
