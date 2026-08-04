! vybe-test: fortran/intent_optional_extended/intent_inout_logical_toggle
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
logical :: flag
flag = .false.
call flip(flag)
if ((flag) .neqv. .true.) then
    print *, "FAIL: want [true] got [", flag, "]"
    stop 1
end if
contains
subroutine flip(v)
logical, intent(inout) :: v
v = .not. v
end subroutine flip
end program t
