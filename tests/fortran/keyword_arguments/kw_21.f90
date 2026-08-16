! vybe-test: fortran/keyword_arguments/kw_21
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
program driver
integer :: seen
seen = 0
call s(1, 2, w=4, z=3)
if (seen /= 1234) then
    print *, "FAIL: want [1234] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x,y,z,w)
integer::x, y, z, w
seen = x * 1000 + y * 100 + z * 10 + w * 1
end
end program driver
