! vybe-test: fortran/arrays/intrinsic_any
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    logical :: mask(3) = [.false., .true., .false.]
    print *, any(mask)
end program test
