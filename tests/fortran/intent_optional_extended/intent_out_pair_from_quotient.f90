! vybe-test: fortran/intent_optional_extended/intent_out_pair_from_quotient
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: q, r
call divide_out(17, 5, q, r)
if ((q) /= 3) then
    print *, "FAIL: want [3] got [", q, "]"
    stop 1
end if
if ((r) /= 2) then
    print *, "FAIL: want [2] got [", r, "]"
    stop 1
end if
contains
subroutine divide_out(n, d, quot, rem)
integer, intent(in) :: n, d
integer, intent(out) :: quot, rem
quot = n / d
rem = mod(n, d)
end subroutine divide_out
end program t
