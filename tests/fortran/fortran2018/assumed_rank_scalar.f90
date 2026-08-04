! vybe-test: fortran/fortran2018/assumed_rank_scalar
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

module ar_mod
    implicit none
contains
    subroutine show_rank(x)
        integer, intent(in) :: x(..)
        print *, rank(x)
    end subroutine show_rank
end module ar_mod

program test
    use ar_mod
    integer :: s = 42
    call show_rank(s)
end program test
