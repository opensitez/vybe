! vybe-test: fortran/program_units/program_optional_args_22
! origin: languages/fortran/tests/fortran/test_program_units.rs
program t
integer :: seen
call s()
if (seen /= -1) then
    print *, "FAIL: want [-1] got [", seen, "]"
    stop 1
end if
call s(8)
if (seen /= 8) then
    print *, "FAIL: want [8] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x)
integer, optional :: x
if (present(x)) then
    seen = x
else
    seen = -1
end if
end subroutine s
end program t
