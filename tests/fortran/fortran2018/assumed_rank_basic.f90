! vybe-test: fortran/fortran2018/assumed_rank_basic
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

module ar_mod
    implicit none
contains
    subroutine describe(x)
        real, intent(in) :: x(..)
        print *, rank(x)
    end subroutine describe
end module ar_mod

program test
    use ar_mod
    real :: a(3,4)
    call describe(a)
end program test
