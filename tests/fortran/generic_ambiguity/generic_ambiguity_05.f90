! vybe-test: fortran/generic_ambiguity/generic_ambiguity_05
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs
module m
interface assignment(=)
module procedure asgi,asgr
end interface
contains
subroutine asgi(a,b)
integer::a,b
a=b
end
subroutine asgr(a,b)
real::a,b
a=b
end
end module m
