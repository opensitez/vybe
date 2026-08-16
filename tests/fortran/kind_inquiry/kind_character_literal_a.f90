! vybe-test: fortran/kind_inquiry/kind_character_literal_a
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((kind('a')) /= 1) then
    print *, "FAIL: want [1] got [", kind('a'), "]"
    stop 1
end if
end program t
