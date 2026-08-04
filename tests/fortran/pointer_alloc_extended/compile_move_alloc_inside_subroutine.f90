! vybe-test: fortran/pointer_alloc_extended/compile_move_alloc_inside_subroutine
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs

subroutine transfer_storage(from, to)
    integer, allocatable, intent(inout) :: from(:)
    integer, allocatable, intent(inout) :: to(:)
    call move_alloc(from, to)
end subroutine transfer_storage

program t
    integer, allocatable :: a(:), b(:)
    allocate(a(2))
    a = [3, 4]
    call transfer_storage(a, b)
    print *, b(2)
    print *, allocated(a)
end program t
