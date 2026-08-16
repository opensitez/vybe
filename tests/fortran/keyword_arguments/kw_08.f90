! vybe-test: fortran/keyword_arguments/kw_08
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
program driver
integer :: seen
seen = 0
call s(j=2,i=1)
if (seen /= 12) then
    print *, "FAIL: want [12] got [", seen, "]"
    stop 1
end if
contains
subroutine s(i,j)
integer::i,j
seen = i * 10 + j * 1
end
end program driver
