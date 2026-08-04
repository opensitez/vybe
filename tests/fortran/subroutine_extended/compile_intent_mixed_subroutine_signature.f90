! vybe-test: fortran/subroutine_extended/compile_intent_mixed_subroutine_signature
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs

program t
    integer :: x, y(2)
    y = [4, 5]
    call mixer(3, x, y)
contains
    subroutine mixer(n, s, arr)
        integer, intent(in) :: n
        integer, intent(out) :: s
        integer, intent(inout) :: arr(2)
        s = n
        arr = arr + n
    end subroutine mixer
end program t
