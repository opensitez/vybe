! vybe-test: fortran/keyword_arguments/kw_03
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
program driver
integer :: seen
seen = 0
call s(z=3,x=1,y=2)
if (seen /= 123) then
    print *, "FAIL: want [123] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x,y,z)
integer::x,y,z
seen = x * 100 + y * 10 + z * 1
end
end program driver
