! vybe-test: fortran/arrays/intrinsic_count
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    logical :: mask(5) = [.true., .false., .true., .true., .false.]
    print *, count(mask)
end program test
