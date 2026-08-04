! vybe-test: fortran/pointer_alloc_extended/associated_with_matching_target_true
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, target :: x = 3, y = 4
integer, pointer :: link
link => x
if ((associated(link, x)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", associated(link, x), "]"
    stop 1
end if
if ((associated(link, y)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", associated(link, y), "]"
    stop 1
end if
end program t
