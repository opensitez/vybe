! vybe-test: fortran/interface_blocks/interface_blocks_16
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
interface operator(.custom.)
integer function custom_add(a, b)
integer, intent(in) :: a, b
end function custom_add
end interface
