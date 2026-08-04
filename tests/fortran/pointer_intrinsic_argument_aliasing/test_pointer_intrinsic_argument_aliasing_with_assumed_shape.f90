! vybe-test: fortran/pointer_intrinsic_argument_aliasing/test_pointer_intrinsic_argument_aliasing_with_assumed_shape
! origin: languages/fortran/tests/fortran/test_pointer_intrinsic_argument_aliasing.rs

program test_pointer_intrinsic_argument_aliasing
    integer, target :: storage
    integer, pointer :: p
    storage = 4
    p => storage
    call mutate(p)
    if ((storage) /= 11) then
    print *, "FAIL: want [11] got [", storage, "]"
    stop 1
end if

contains
    subroutine mutate(value)
        integer, pointer, intent(inout) :: value
        value = 11
    end subroutine
end program test_pointer_intrinsic_argument_aliasing
