! vybe-test: fortran/pointer_alloc_extended/pointer_write_updates_target_storage
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, target :: base = 1
integer, pointer :: alias
alias => base
alias = 50
if ((base) /= 50) then
    print *, "FAIL: want [50] got [", base, "]"
    stop 1
end if
end program t
