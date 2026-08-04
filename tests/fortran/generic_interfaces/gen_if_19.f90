! vybe-test: fortran/generic_interfaces/gen_if_19
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface g
module procedure fs
end interface
contains
integer function fs()
fs=1
end
end module m
