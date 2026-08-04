! vybe-test: fortran/forall_advanced/forall_in_subroutine
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    real :: a(5)
    call fill_squares(a)
    print *, a(3)
contains
    subroutine fill_squares(x)
        real, intent(out) :: x(:)
        integer :: n
        n = size(x)
        forall (i = 1:n)
            x(i) = real(i) ** 2
        end forall
    end subroutine fill_squares
end program test
