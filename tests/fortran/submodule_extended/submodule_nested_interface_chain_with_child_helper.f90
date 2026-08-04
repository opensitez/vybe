! vybe-test: fortran/submodule_extended/submodule_nested_interface_chain_with_child_helper
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module chain_iface
    implicit none
    interface
        module function top(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function top
    end interface
end module chain_iface

submodule (chain_iface) chain_mid
    interface
        module function helper(x) result(r)
            integer, intent(in) :: x
            integer :: r
        end function helper
    end interface
end submodule chain_mid

submodule (chain_iface:chain_mid) chain_leaf
    contains
    module function top(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = helper(x) + 3
    end function top

    module function helper(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * 2
    end function helper
end submodule chain_leaf

program t
    use chain_iface
    if ((top(4)) /= 11) then
    print *, "FAIL: want [11] got [", top(4), "]"
    stop 1
end if
end program t
