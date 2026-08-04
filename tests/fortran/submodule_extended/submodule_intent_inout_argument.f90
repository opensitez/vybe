! vybe-test: fortran/submodule_extended/submodule_intent_inout_argument
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module io_iface
    implicit none
    interface
        module subroutine double_it(x)
            integer, intent(inout) :: x
        end subroutine double_it
    end interface
end module io_iface

submodule (io_iface) io_impl
contains
    module subroutine double_it(x)
        integer, intent(inout) :: x
        x = x * 2
    end subroutine double_it
end submodule io_impl

program t
    use io_iface
    integer :: n = 6
    call double_it(n)
    if ((n) /= 12) then
    print *, "FAIL: want [12] got [", n, "]"
    stop 1
end if
end program t
