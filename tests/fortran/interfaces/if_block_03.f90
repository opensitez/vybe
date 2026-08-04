! vybe-test: fortran/interfaces/if_block_03
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
interface
subroutine s(x)
integer :: x
end subroutine s
end interface
end module m
