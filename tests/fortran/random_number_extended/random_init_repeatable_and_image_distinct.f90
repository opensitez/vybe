! vybe-test: fortran/random_number_extended/random_init_repeatable_and_image_distinct
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    call random_init(repeatable=.true., image_distinct=.true.)
    real :: r
    call random_number(r)
    print *, r >= 0.0
end program t
