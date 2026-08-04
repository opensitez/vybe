! vybe-test: fortran/arrays_shift/cshift_1d_logical
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    logical :: a(4) = [.true., .false., .true., .false.]
    logical :: b(4)
    b = cshift(a, 1)
    print *, b(1)
end program test
