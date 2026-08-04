! vybe-test: fortran/interface_blocks/interface_blocks_11
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
interface
subroutine s(x, y)
integer, intent(inout) :: x
integer, intent(in) :: y
end subroutine s
end interface
