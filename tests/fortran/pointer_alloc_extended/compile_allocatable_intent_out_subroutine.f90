! vybe-test: fortran/pointer_alloc_extended/compile_allocatable_intent_out_subroutine
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs

subroutine make_buffer(buf, n)
    integer, intent(in) :: n
    integer, allocatable, intent(out) :: buf(:)
    allocate(buf(n))
    buf = [(i, i = 1, n)]
end subroutine make_buffer

program t
    integer, allocatable :: data(:)
    call make_buffer(data, 3)
    print *, sum(data)
end program t
