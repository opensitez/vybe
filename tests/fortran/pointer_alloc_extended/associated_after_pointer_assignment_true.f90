! vybe-test: fortran/pointer_alloc_extended/associated_after_pointer_assignment_true
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, target :: host = 17
integer, pointer :: view
view => host
if ((associated(view)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", associated(view), "]"
    stop 1
end if
end program t
