! vybe-test: fortran/program_units/program_target_arg_28
! origin: languages/fortran/tests/fortran/test_program_units.rs
program t
integer :: v
v = 0
call s(v)
if (v /= 4) then
    print *, "FAIL: want [4] got [", v, "]"
    stop 1
end if
contains
subroutine s(a)
integer, target :: a
integer, pointer :: p
p => a
p = 4
end subroutine s
end program t
