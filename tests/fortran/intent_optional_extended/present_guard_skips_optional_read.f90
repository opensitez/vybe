! vybe-test: fortran/intent_optional_extended/present_guard_skips_optional_read
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((guarded_add(5)) /= 5) then
    print *, "FAIL: want [5] got [", guarded_add(5), "]"
    stop 1
end if
if ((guarded_add(5, 8)) /= 13) then
    print *, "FAIL: want [13] got [", guarded_add(5, 8), "]"
    stop 1
end if
contains
integer function guarded_add(a, b)
integer, intent(in) :: a
integer, intent(in), optional :: b
guarded_add = a
if (present(b)) guarded_add = guarded_add + b
end function guarded_add
end program t
