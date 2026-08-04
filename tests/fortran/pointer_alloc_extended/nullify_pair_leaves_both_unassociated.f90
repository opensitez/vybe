! vybe-test: fortran/pointer_alloc_extended/nullify_pair_leaves_both_unassociated
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, target :: u = 1, v = 2
integer, pointer :: p, q
p => u
q => v
nullify(p, q)
if ((associated(p)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", associated(p), "]"
    stop 1
end if
if ((associated(q)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", associated(q), "]"
    stop 1
end if
end program t
