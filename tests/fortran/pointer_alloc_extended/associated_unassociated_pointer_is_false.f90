! vybe-test: fortran/pointer_alloc_extended/associated_unassociated_pointer_is_false
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, pointer :: p => null()
if ((associated(p)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", associated(p), "]"
    stop 1
end if
end program t
