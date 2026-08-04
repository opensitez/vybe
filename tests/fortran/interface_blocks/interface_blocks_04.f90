! vybe-test: fortran/interface_blocks/interface_blocks_04
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
module m
interface
subroutine s(x)
integer::x
end subroutine s
end interface
end module m
