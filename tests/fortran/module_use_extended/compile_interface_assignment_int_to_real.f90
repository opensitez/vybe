! vybe-test: fortran/module_use_extended/compile_interface_assignment_int_to_real
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs

module assign_iface
    implicit none
    interface assignment(=)
        module procedure int_assign_real
    end interface
contains
    subroutine int_assign_real(r, i)
        real, intent(out) :: r
        integer, intent(in) :: i
        r = real(i)
    end subroutine int_assign_real
end module assign_iface

program t
    use assign_iface
    real :: x
    x = 5
    print *, x
end program t
