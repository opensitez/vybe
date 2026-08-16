! vybe-test: fortran/intent_attributes/intent_attributes_08
! origin: languages/fortran/tests/fortran/test_intent_attributes.rs
program t
logical :: flag
flag = .false.
call s(flag)
if (.not. flag) then
    print *, "FAIL: want [true] got [", flag, "]"
    stop 1
end if
contains
subroutine s(x)
logical, intent(out) :: x
x = .true.
end subroutine s
end program t
