! vybe-test: fortran/interface_blocks/interface_blocks_17
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
interface assignment(=)
subroutine assign_wrapper(lhs, rhs)
integer, intent(out) :: lhs
integer, intent(in) :: rhs
end subroutine assign_wrapper
end interface
