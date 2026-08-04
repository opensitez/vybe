! vybe-test: fortran/interfaces/if_generic_optional_47
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
interface g
module procedure g_int, g_real
end interface
contains
integer function g_int(x, flag)
integer, intent(in) :: x
integer, intent(in), optional :: flag
g_int = x
end function g_int
real function g_real(x, flag)
real, intent(in) :: x
logical, intent(in), optional :: flag
g_real = x
end function g_real
end module m
