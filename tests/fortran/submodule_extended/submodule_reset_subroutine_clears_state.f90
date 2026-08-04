! vybe-test: fortran/submodule_extended/submodule_reset_subroutine_clears_state
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module acc_iface
    implicit none
    interface
        module subroutine add_one()
        end subroutine add_one
        module subroutine reset_acc()
        end subroutine reset_acc
        module function read_acc() result(n)
            integer :: n
        end function read_acc
    end interface
end module acc_iface

submodule (acc_iface) acc_impl
    integer :: total = 0
contains
    module subroutine add_one()
        total = total + 1
    end subroutine add_one

    module subroutine reset_acc()
        total = 0
    end subroutine reset_acc

    module function read_acc() result(n)
        integer :: n
        n = total
    end function read_acc
end submodule acc_impl

program t
    use acc_iface
    call add_one()
    call add_one()
    call reset_acc()
    call add_one()
    if ((read_acc()) /= 1) then
    print *, "FAIL: want [1] got [", read_acc(), "]"
    stop 1
end if
end program t
