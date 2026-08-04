! vybe-test: fortran/intent_optional_extended/intent_out_rank_two_shape_init
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
integer :: m(2, 2)
call identity2(m)
if ((sum(m)) /= 2) then
    print *, "FAIL: want [2] got [", sum(m), "]"
    stop 1
end if
contains
subroutine identity2(a)
integer, intent(out) :: a(2, 2)
a = 0
a(1, 1) = 1
a(2, 2) = 1
end subroutine identity2
end program t
