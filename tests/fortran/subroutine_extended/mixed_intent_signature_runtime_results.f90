! vybe-test: fortran/subroutine_extended/mixed_intent_signature_runtime_results
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs

program t
    integer :: x
    integer :: y(2)
    y = [4, 5]
    call mixer(3, x, y)
    if ((x) /= 3) then
    print *, "FAIL: want [3] got [", x, "]"
    stop 1
end if
    if ((sum(y)) /= 15) then
    print *, "FAIL: want [15] got [", sum(y), "]"
    stop 1
end if
contains
    subroutine mixer(n, s, arr)
        integer, intent(in) :: n
        integer, intent(out) :: s
        integer, intent(inout) :: arr(2)
        s = n
        arr = arr + n
    end subroutine mixer
end program t
