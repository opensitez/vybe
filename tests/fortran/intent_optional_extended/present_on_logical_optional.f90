! vybe-test: fortran/intent_optional_extended/present_on_logical_optional
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((opt_and(.true.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", opt_and(.true.), "]"
    stop 1
end if
if ((opt_and(.true., .false.)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", opt_and(.true., .false.), "]"
    stop 1
end if
contains
logical function opt_and(a, b)
logical, intent(in) :: a
logical, intent(in), optional :: b
if (present(b)) then
opt_and = a .and. b
else
opt_and = a
end if
end function opt_and
end program t
