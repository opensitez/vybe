! vybe-test: fortran/pointer_alloc_extended/pointer_read_scalar_through_association
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, target :: base = 42
integer, pointer :: alias
alias => base
if ((alias) /= 42) then
    print *, "FAIL: want [42] got [", alias, "]"
    stop 1
end if
end program t
