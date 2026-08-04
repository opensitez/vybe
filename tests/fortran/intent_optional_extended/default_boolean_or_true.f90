! vybe-test: fortran/intent_optional_extended/default_boolean_or_true
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((coalesce_bool(.false.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", coalesce_bool(.false.), "]"
    stop 1
end if
if ((coalesce_bool(.false., .true.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", coalesce_bool(.false., .true.), "]"
    stop 1
end if
contains
logical function coalesce_bool(v, alt)
logical, intent(in) :: v
logical, intent(in), optional :: alt
if (present(alt)) then
coalesce_bool = alt
else
coalesce_bool = .true.
end if
end function coalesce_bool
end program t
