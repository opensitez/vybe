! vybe-test: fortran/generic_interfaces/gen_if_12
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface operator(//)
module procedure cat
end interface
contains
character(len=2) function cat(a,b)
character(len=1)::a,b
cat=a//b
end
end module m
