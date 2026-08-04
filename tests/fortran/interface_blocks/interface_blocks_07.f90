! vybe-test: fortran/interface_blocks/interface_blocks_07
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
interface
subroutine s(a)
real::a(:)
end subroutine s
end interface
