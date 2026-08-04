! vybe-test: fortran/submodule_extended/submodule_child_interface_grandchild_implements
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module top_iface
    implicit none
    interface
        module function top_val() result(r)
            integer :: r
        end function top_val
    end interface
end module top_iface

submodule (top_iface) mid_iface
    implicit none
    interface
        module function mid_val() result(r)
            integer :: r
        end function mid_val
    end interface
end submodule mid_iface

submodule (top_iface:mid_iface) bot_impl
    implicit none
contains
    module function top_val() result(r)
        integer :: r
        r = 10
    end function top_val

    module function mid_val() result(r)
        integer :: r
        r = 20
    end function mid_val
end submodule bot_impl

program t
    use top_iface
    if ((top_val()) /= 10) then
    print *, "FAIL: want [10] got [", top_val(), "]"
    stop 1
end if
end program t
