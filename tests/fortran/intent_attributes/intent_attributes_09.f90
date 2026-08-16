! vybe-test: fortran/intent_attributes/intent_attributes_09
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program t
complex :: z
z = (1.0, 2.0)
call s(z)
if (abs(real(z) - 2.0) > 1.0e-6) then
    print *, "FAIL: want [2.0] got [", real(z), "]"
    stop 1
end if
if (abs(aimag(z) - 4.0) > 1.0e-6) then
    print *, "FAIL: want [4.0] got [", aimag(z), "]"
    stop 1
end if
contains
subroutine s(x)
complex, intent(inout) :: x
x = x * 2.0
end subroutine s
end program t
