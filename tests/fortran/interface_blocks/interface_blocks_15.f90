! vybe-test: fortran/interface_blocks/interface_blocks_15
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
interface
character(len=4) function aschar(i)
integer, intent(in) :: i
end function aschar
end interface
