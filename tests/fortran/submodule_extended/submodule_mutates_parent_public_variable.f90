! vybe-test: fortran/submodule_extended/submodule_mutates_parent_public_variable
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module bank_iface
    implicit none
    integer :: balance = 100
    interface
        module subroutine deposit(amt)
            integer, intent(in) :: amt
        end subroutine deposit
    end interface
end module bank_iface

submodule (bank_iface) bank_impl
contains
    module subroutine deposit(amt)
        integer, intent(in) :: amt
        balance = balance + amt
    end subroutine deposit
end submodule bank_impl

program t
    use bank_iface
    call deposit(50)
    if ((balance) /= 150) then
    print *, "FAIL: want [150] got [", balance, "]"
    stop 1
end if
end program t
