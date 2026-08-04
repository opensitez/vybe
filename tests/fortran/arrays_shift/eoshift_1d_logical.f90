! vybe-test: fortran/arrays_shift/eoshift_1d_logical
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    logical :: a(4) = [.true., .true., .false., .true.]
    logical :: b(4)
    b = eoshift(a, -1, boundary=.false.)
    print *, b(1)
end program test
