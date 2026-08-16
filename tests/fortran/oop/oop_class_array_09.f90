! vybe-test: fortran/oop/oop_class_array_09
! origin: languages/fortran/tests/fortran/test_oop.rs
program p
type::t
integer::x
end type t
class(t), allocatable :: a(:)
allocate(a(2))
end program p
