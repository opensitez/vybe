! vybe-test: fortran/pointer_alloc_extended/nullify_then_reassociate_pointer
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, target :: a = 1, b = 2
integer, pointer :: p
p => a
nullify(p)
if ((associated(p)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", associated(p), "]"
    stop 1
end if
p => b
if ((p) /= 2) then
    print *, "FAIL: want [2] got [", p, "]"
    stop 1
end if
end program t
