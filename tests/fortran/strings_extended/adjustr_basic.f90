! vybe-test: fortran/strings_extended/adjustr_basic
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=10) :: s = 'hello'
character(len=10) :: r
r = adjustr(s)
if ((len_trim(r)) /= 5) then
    print *, "FAIL: want [5] got [", len_trim(r), "]"
    stop 1
end if
end program t
