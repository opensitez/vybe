! vybe-test: fortran/pointer_alloc_extended/compile_pointer_and_allocatable_in_module
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs

module storage
    implicit none
    type :: Slot
        integer, pointer :: view => null()
        integer, allocatable :: owned(:)
    end type Slot
contains
    subroutine attach(slot, target_arr)
        type(Slot), intent(inout) :: slot
        integer, target, intent(in) :: target_arr(:)
        slot%view => target_arr
        if (.not. allocated(slot%owned)) allocate(slot%owned(size(target_arr)))
        slot%owned = target_arr
    end subroutine attach
end module storage

program t
    use storage
    integer, target :: data(2) = [11, 22]
    type(Slot) :: box
    call attach(box, data)
    print *, box%view(2)
    print *, box%owned(1)
end program t
