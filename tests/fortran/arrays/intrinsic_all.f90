! vybe-test: fortran/arrays/intrinsic_all
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    logical :: mask(3) = [.true., .true., .true.]
    print *, all(mask)
end program test
