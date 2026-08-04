! vybe-test: fortran/subroutine_extended/intent_out_two_outputs_from_sub
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
integer :: p, q
call split_sum(10, p, q)
if ((p) /= 5) then
    print *, "FAIL: want [5] got [", p, "]"
    stop 1
end if
if ((q) /= 5) then
    print *, "FAIL: want [5] got [", q, "]"
    stop 1
end if
contains
subroutine split_sum(n, half, rest)
integer, intent(in) :: n
integer, intent(out) :: half, rest
half = n / 2
rest = n - half
end subroutine split_sum
end program t
