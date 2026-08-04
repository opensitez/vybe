! vybe-test: fortran/modules_advanced/interface_assignment
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module conv
    implicit none
    interface assignment(=)
        module procedure int_to_real
    end interface
contains
    subroutine int_to_real(r, i)
        real, intent(out) :: r
        integer, intent(in) :: i
        r = real(i)
    end subroutine int_to_real
end module conv

program test
    print *, "ok"
end program test
