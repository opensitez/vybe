! vybe-test: fortran/kind_inquiry/kind_character_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
character(len=12) :: s
s = 'abc'
if ((kind(s)) /= 8) then
    print *, "FAIL: want [8] got [", kind(s), "]"
    stop 1
end if
end program t
