! vybe-test: fortran/pointer_alloc_extended/associated_after_nullify_is_false
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, target :: val = 8
integer, pointer :: link
link => val
nullify(link)
if ((associated(link)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", associated(link), "]"
    stop 1
end if
end program t
