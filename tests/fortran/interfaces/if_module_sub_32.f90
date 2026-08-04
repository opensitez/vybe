! vybe-test: fortran/interfaces/if_module_sub_32
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
interface
module subroutine s(x)
integer :: x
end subroutine s
end interface
end module m
