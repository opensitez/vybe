! vybe-test: fortran/generic_interfaces/gen_if_07
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface operator(*)
module procedure muli
end interface
contains
integer function muli(a,b)
integer::a,b
muli=a*b
end
end module m
