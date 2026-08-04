! vybe-test: fortran/generic_interfaces/gen_if_01
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface g
module procedure s1
end interface
contains
subroutine s1()
end subroutine s1
end module m
