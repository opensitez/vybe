! vybe-test: fortran/keyword_arguments/kw_15
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
program driver
integer :: seen
seen = 0
call s(b=2, a=1)
if (seen /= 12) then
    print *, "FAIL: want [12] got [", seen, "]"
    stop 1
end if
contains
subroutine s(a,b)
integer::a,b
seen = a * 10 + b * 1
end
end program driver
