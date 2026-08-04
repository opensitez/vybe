! vybe-test: fortran/fortran2018_extended/random_init_in_module_initializer
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

module rng
    implicit none
contains
    subroutine seed_once()
        call random_init(repeatable=.false., image_distinct=.false.)
    end subroutine seed_once
end module rng

program t
    use rng
    call seed_once()
    print *, 'seeded'
end program t
