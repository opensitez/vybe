! vybe-test: fortran/pointers/pointer_dummy_allocation_preserves_null_entry
! origin: languages/fortran/tests/fortran/test_pointers.rs

module m
    implicit none

    type :: node
        integer :: value = 0
        type(node), pointer :: next => null()
    end type node
contains
    subroutine ensure_value(item, value)
        type(node), pointer, intent(inout) :: item
        integer, intent(in) :: value

        if (.not. associated(item)) then
            allocate(item)
            item%value = value
            nullify(item%next)
        end if
    end subroutine ensure_value
end module m

program test
    use m
    implicit none
    type(node), pointer :: head => null()

    call ensure_value(head, 5)
    if ((associated(head)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", associated(head), "]"
    stop 1
end if
    if (associated(head)) then
        if ((head%value) /= 5) then
    print *, "FAIL: want [5] got [", head%value, "]"
    stop 1
end if
    end if
end program test
