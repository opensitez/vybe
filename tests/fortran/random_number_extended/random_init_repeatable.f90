! vybe-test: fortran/random_number_extended/random_init_repeatable
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    call random_init(repeatable=.true., image_distinct=.false.)
    real :: x
    call random_number(x)
    print *, x >= 0.0 .and. x < 1.0
end program test
