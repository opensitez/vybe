! vybe-test: fortran/submodule_extended/submodule_parent_var_initial_before_sub_call
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
if ((balance) /= 100) then
    print *, "FAIL: want [100] got [", balance, "]"
    stop 1
end if
end program t
