! vybe-test: fortran/generic_resolution/generic_resolution_09
! origin: languages/fortran/tests/fortran/test_generic_resolution.rs
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
