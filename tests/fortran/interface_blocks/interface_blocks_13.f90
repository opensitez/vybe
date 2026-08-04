! vybe-test: fortran/interface_blocks/interface_blocks_13
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
interface
subroutine copy(src, dst)
integer, intent(in) :: src(:)
integer, intent(out) :: dst(:)
end subroutine copy
end interface
