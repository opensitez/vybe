! vybe-test: fortran/submodules_advanced/submodule_with_protected_parent_var_runtime
! origin: languages/fortran/tests/fortran/test_submodules_advanced.rs

module cfg_iface
    implicit none
    integer, protected :: max_size = 100
    interface
        module subroutine set_max(n)
            integer, intent(in) :: n
        end subroutine set_max
    end interface
end module cfg_iface

submodule (cfg_iface) cfg_impl
    implicit none
contains
    module subroutine set_max(n)
        integer, intent(in) :: n
        max_size = n
    end subroutine set_max
end submodule cfg_impl

program test
    use cfg_iface
    if ((max_size) /= 100) then
    print *, "FAIL: want [100] got [", max_size, "]"
    stop 1
end if
    call set_max(200)
    if ((max_size) /= 200) then
    print *, "FAIL: want [200] got [", max_size, "]"
    stop 1
end if
end program test
