! vybe-test: fortran/interface_blocks/interface_blocks_14
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
interface
logical function has_value(v)
integer, intent(in) :: v
end function has_value
end interface
