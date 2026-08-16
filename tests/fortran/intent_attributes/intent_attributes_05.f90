! vybe-test: fortran/intent_attributes/intent_attributes_05
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program t
real :: buf(3)
buf = [9.0, 9.0, 9.0]
call s(buf)
if (abs(sum(buf) - 3.0) > 1.0e-6) then
    print *, "FAIL: want [3.0] got [", sum(buf), "]"
    stop 1
end if
contains
subroutine s(a)
real, intent(out) :: a(:)
a = 1.0
end subroutine s
end program t
