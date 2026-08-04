! vybe-test: fortran/allocation/assumed_type_12
! origin: languages/fortran/tests/fortran/test_allocation.rs
subroutine s(x)
type(*) :: x
end subroutine s
