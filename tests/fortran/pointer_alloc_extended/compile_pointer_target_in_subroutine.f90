! vybe-test: fortran/pointer_alloc_extended/compile_pointer_target_in_subroutine
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs

subroutine bind_ptr(host, view)
    integer, target, intent(inout) :: host
    integer, pointer, intent(out) :: view
    view => host
end subroutine bind_ptr

program t
    integer, target :: x = 6
    integer, pointer :: p
    call bind_ptr(x, p)
    print *, p
end program t
