! vybe-test: fortran/intent_attributes/intent_attributes_04
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program t
real :: buf(3)
real :: total
buf = [1.0, 2.0, 3.0]
call s(buf)
if (abs(total - 6.0) > 1.0e-6) then
    print *, "FAIL: want [6.0] got [", total, "]"
    stop 1
end if
contains
subroutine s(a)
real, intent(in) :: a(:)
total = sum(a)
end subroutine s
end program t
