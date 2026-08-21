! vybe-test: fortran/random_number_extended/random_init_non_repeatable
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    call random_init(repeatable=.false., image_distinct=.true.)
    real :: r
    call random_number(r)
    print *, 'ok'
end program test
