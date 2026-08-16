! vybe-test: fortran/program_units/program_keyword_args_23
! origin: languages/fortran/tests/fortran/test_program_units.rs
program driver
integer :: seen
seen = 0
call s(y=2, x=1)
if (seen /= 12) then
    print *, "FAIL: want [12] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x,y)
integer :: x,y
seen = x * 10 + y * 1
end subroutine s
end program driver
