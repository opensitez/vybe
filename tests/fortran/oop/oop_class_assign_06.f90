! vybe-test: fortran/oop/oop_class_assign_06
! origin: languages/fortran/tests/fortran/test_oop.rs
type::t
integer::x
end type t
program p
class(t),allocatable::a,b
allocate(a,b)
a=b
end program p
