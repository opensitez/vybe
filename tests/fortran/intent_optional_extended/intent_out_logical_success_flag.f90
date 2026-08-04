! vybe-test: fortran/intent_optional_extended/intent_out_logical_success_flag
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
logical :: ok
call set_ok(ok)
if ((ok) .neqv. .true.) then
    print *, "FAIL: want [true] got [", ok, "]"
    stop 1
end if
contains
subroutine set_ok(flag)
logical, intent(out) :: flag
flag = .true.
end subroutine set_ok
end program t
