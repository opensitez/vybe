! vybe-test: fortran/generic_interfaces/gen_if_20
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface operator(==)
module procedure eqi
end interface
contains
logical function eqi(a,b)
integer::a,b
eqi=a==b
end
end module m
