! vybe-test: fortran/generic_interfaces/gen_if_16
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs
module m
interface operator(.foo.)
module procedure foo
end interface
contains
logical function foo(a,b)
logical::a,b
foo=a.and.b
end
end module m
