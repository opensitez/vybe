! vybe-test: fortran/pointer_alloc_extended/compile_pointer_deferred_shape_dummy
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs

subroutine first_elt(vec, out)
    integer, pointer, intent(in) :: vec(:)
    integer, intent(out) :: out
    out = vec(1)
end subroutine first_elt

program t
    integer, target :: arr(4) = [8, 6, 4, 2]
    integer, pointer :: p(:)
    integer :: head
    p => arr
    call first_elt(p, head)
    print *, head
end program t
