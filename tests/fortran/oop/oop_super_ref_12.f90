! vybe-test: fortran/oop/oop_super_ref_12
! origin: languages/fortran/tests/fortran/test_oop.rs
module m
type::b
contains
procedure::show
end type b
type,extends(b)::c
contains
procedure::show=>show_c
end type c
contains
subroutine show(this)
class(b)::this
end
subroutine show_c(this)
class(c)::this
end
end module m
