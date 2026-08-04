! vybe-test: fortran/fortran2018/assumed_rank_2d
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

module ar_mod
    implicit none
contains
    subroutine show(a)
        integer, intent(in) :: a(..)
        select rank(a)
        rank(2)
            print *, size(a,1), size(a,2)
        rank default
            print *, rank(a)
        end select
    end subroutine show
end module ar_mod

program test
    use ar_mod
    integer :: m(4,4)
    call show(m)
end program test
