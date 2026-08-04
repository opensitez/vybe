! vybe-test: fortran/pure_elemental/pure_subroutine
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    integer :: x = 5
    call double_it(x)
    print *, x
contains
    pure subroutine double_it(n)
        integer, intent(inout) :: n
        n = n * 2
    end subroutine double_it
end program test
