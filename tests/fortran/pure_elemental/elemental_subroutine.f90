! vybe-test: fortran/pure_elemental/elemental_subroutine
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    integer :: a(3) = [1, 2, 3]
    call negate(a)
    print *, a(1)
contains
    elemental subroutine negate(x)
        integer, intent(inout) :: x
        x = -x
    end subroutine negate
end program test
