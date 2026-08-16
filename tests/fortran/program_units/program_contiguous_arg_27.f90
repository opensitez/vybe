! vybe-test: fortran/program_units/program_contiguous_arg_27
! origin: languages/fortran/tests/fortran/test_program_units.rs
program t
real :: buf(3)
buf = [1.0, 2.0, 3.0]
call s(buf)
if (abs(sum(buf) - 12.0) > 1.0e-6) then
    print *, "FAIL: want [12.0] got [", sum(buf), "]"
    stop 1
end if
contains
subroutine s(a)
real, contiguous :: a(:)
a = a * 2.0
end subroutine s
end program t
