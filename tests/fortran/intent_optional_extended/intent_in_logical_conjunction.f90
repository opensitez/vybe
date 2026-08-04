! vybe-test: fortran/intent_optional_extended/intent_in_logical_conjunction
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((land(.true., .false.)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", land(.true., .false.), "]"
    stop 1
end if
if ((land(.true., .true.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", land(.true., .true.), "]"
    stop 1
end if
contains
logical function land(p, q)
logical, intent(in) :: p, q
land = p .and. q
end function land
end program t
