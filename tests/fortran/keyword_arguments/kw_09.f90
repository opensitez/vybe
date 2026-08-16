! vybe-test: fortran/keyword_arguments/kw_09
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
program driver
integer :: seen
seen = 0
call s(x=1)
if (seen /= 1) then
    print *, "FAIL: want [1] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x)
integer::x
seen = x * 1
end
end program driver
