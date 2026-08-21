! vybe-test: fortran/reduce_intrinsic/reduce_mask
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: a(6) = [1, 2, 3, 4, 5, 6]
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    integer :: r
    r = reduce(a, operator(+), mask=mask)
    print *, r
end program test
