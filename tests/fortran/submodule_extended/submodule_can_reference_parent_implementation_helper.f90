! vybe-test: fortran/submodule_extended/submodule_can_reference_parent_implementation_helper
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module base_iface
    implicit none
contains
    integer function offset_base()
        offset_base = 20
    end function offset_base
    interface
        module function apply_offset(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function apply_offset
    end interface
end module base_iface

submodule (base_iface) base_impl
contains
    module function apply_offset(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x + offset_base()
    end function apply_offset
end submodule base_impl

program t
    use base_iface
    if ((apply_offset(3)) /= 23) then
    print *, "FAIL: want [23] got [", apply_offset(3), "]"
    stop 1
end if
end program t
