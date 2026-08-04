! vybe-test: fortran/allocation_semantics/as_14
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
type :: t
character(len=:), allocatable :: s
end type t
program p
type(t) :: x
allocate(character(len=3) :: x%s)
end program p
