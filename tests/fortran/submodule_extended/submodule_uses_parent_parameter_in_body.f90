! vybe-test: fortran/submodule_extended/submodule_uses_parent_parameter_in_body
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module scale_iface
    implicit none
    integer, parameter :: SCALE = 10
    interface
        module function scaled(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function scaled
    end interface
end module scale_iface

submodule (scale_iface) scale_impl
    implicit none
contains
    module function scaled(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * SCALE
    end function scaled
end submodule scale_impl

program t
    use scale_iface
    if ((scaled(4)) /= 40) then
    print *, "FAIL: want [40] got [", scaled(4), "]"
    stop 1
end if
end program t
