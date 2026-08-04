! vybe-test: fortran/interface_blocks/interface_blocks_12
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
interface
integer function f(a, b, scale)
integer, intent(in) :: a, b
integer, intent(in), optional :: scale
end function f
end interface
