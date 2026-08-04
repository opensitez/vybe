! vybe-test: fortran/interface_blocks/interface_blocks_20
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
interface
subroutine s(x)
integer, optional, intent(in) :: x
end subroutine s
end interface
