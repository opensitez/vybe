! vybe-test: fortran/interfaces/if_generic_04
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
interface g
module procedure s1
end interface
contains
subroutine s1()
end subroutine s1
end module m
