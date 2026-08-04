! vybe-test: fortran/submodules_advanced/submodule_private_state_runtime
! origin: languages/fortran/tests/fortran/test_submodules_advanced.rs

module counter_iface
    implicit none
    interface
        module subroutine increment()
        end subroutine increment
        module function get_count() result(n)
            integer :: n
        end function get_count
    end interface
end module counter_iface

submodule (counter_iface) counter_impl
    implicit none
    integer :: count = 0
contains
    module subroutine increment()
        count = count + 1
    end subroutine increment

    module function get_count() result(n)
        integer :: n
        n = count
    end function get_count
end submodule counter_impl

program test
    use counter_iface
    call increment()
    call increment()
    call increment()
    if ((get_count()) /= 3) then
    print *, "FAIL: want [3] got [", get_count(), "]"
    stop 1
end if
end program test
