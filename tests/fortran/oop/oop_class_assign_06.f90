! vybe-test: fortran/oop/oop_class_assign_06
! origin: languages/fortran/tests/fortran/test_oop.rs
program p
type::t
integer::x
end type t
class(t),allocatable::a,b
allocate(a,b)
a=b
end program p
