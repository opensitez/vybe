! vybe-test: fortran/allocation_semantics/as_14
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
type :: t
character(len=:), allocatable :: s
end type t
type(t) :: x
allocate(character(len=3) :: x%s)
end program p
